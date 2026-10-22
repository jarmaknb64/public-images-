/**
 * Intelligems Campaign Bot — v1.1
 *
 * Pattern: same as Survey Bot and Celebrations Bot.
 *   - Single Google Sheet as database
 *   - Hourly orchestrator()
 *   - Slack reaction_added events hit doPost()
 *
 * Setup checklist:
 *   1. Sheet tabs required: Campaigns, Recipients, Campaign_Log, HRIS_Roster, Config
 *   2. Apps Script Script Properties needed:
 *        SLACK_TOKEN              xoxp user token
 *        SLACK_SIGNING_SECRET     from Slack app config (used for v2 webhook verification)
 *        HRIS_SHEET_ID            this sheet ID (or external roster sheet ID)
 *        SURVEY_BOT_SHEET_ID      this sheet ID (or wherever Survey Bot's Audience tab lives)
 *        DEFAULT_FOLLOWUP_CHANNEL Slack channel ID for 🙋 routing
 *   3. Slack app scopes required (User Token Scopes for the xoxp token):
 *        chat:write, im:write, users:read, reactions:read, reactions:write,
 *        conversations:read, groups:read
 *   4. Slack Events API: subscribe to reaction_added and reaction_removed (User events)
 *   5. Deploy as web app: Execute as me, Anyone can access
 *   6. Set hourly trigger on orchestrator()
 *
 * HRIS_Roster schema (slim version):
 *   name | slack_user_id | leadership
 *
 * Audience filter syntax:
 *   all_employees                      Everyone in HRIS_Roster
 *   leadership:true                    People with leadership=TRUE
 *   slack_channel:C12345               Everyone in that Slack channel
 *   manual:U012,U345                   Specific Slack user IDs, comma-separated
 *   <filter> AND <filter>              Compound, e.g. leadership:true AND slack_channel:C123
 */

// ─────────────────────────────────────────────────────────────────────────────
// CONFIG
// ─────────────────────────────────────────────────────────────────────────────

const SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const PROPS = PropertiesService.getScriptProperties();
const SLACK_TOKEN = PROPS.getProperty('SLACK_TOKEN');
const SLACK_SIGNING_SECRET = PROPS.getProperty('SLACK_SIGNING_SECRET');
const HRIS_SHEET_ID = PROPS.getProperty('HRIS_SHEET_ID') || SHEET_ID;
const SURVEY_BOT_SHEET_ID = PROPS.getProperty('SURVEY_BOT_SHEET_ID') || SHEET_ID;
const DEFAULT_FOLLOWUP_CHANNEL = PROPS.getProperty('DEFAULT_FOLLOWUP_CHANNEL');

const TABS = {
  CAMPAIGNS: 'Campaigns',
  RECIPIENTS: 'Recipients',
  LOG: 'Campaign_Log',
  ROSTER: 'HRIS_Roster',
  CONFIG: 'Config',
};

const SEND_DELAY_MS = 1000; // 1s between bulk sends to stay under Slack rate limits

// ─────────────────────────────────────────────────────────────────────────────
// MAIN ENTRY POINT
// ─────────────────────────────────────────────────────────────────────────────

function orchestrator() {
  console.log('=== orchestrator() started ===');

  const roster = getHrisRoster();
  console.log('Roster size: ' + roster.length);
  if (roster.length > 0) {
    console.log('First roster entry: ' + JSON.stringify(roster[0]));
  }

  const killSwitch = getConfigValue('kill_switch');
  console.log('Kill switch value: ' + killSwitch);
  if (killSwitch === 'TRUE') {
    console.log('Kill switch is TRUE, exiting');
    log('orchestrator', 'KILL SWITCH ACTIVE — exiting');
    return;
  }

  const campaigns = readTab(TABS.CAMPAIGNS);
  console.log('Found ' + campaigns.length + ' campaigns total');

  for (const c of campaigns) {
    console.log('---');
    console.log('Campaign: ' + c.campaign_id + ' | status: ' + c.status);

    try {
      if (c.status === 'picker_open') {
        console.log('  Status is picker_open, attempting enrollment');
        console.log('  Audience filter: "' + c.audience_filter + '"');
        const eligible = resolveAudienceFilter(c.audience_filter, c);
        console.log('  Filter resolved to ' + eligible.length + ' people');
        if (eligible.length > 0) {
          console.log('  Eligible IDs: ' + eligible.map(e => e.slack_user_id).join(', '));
        }
        const existing = getRecipientsForCampaign(c.campaign_id);
        console.log('  Existing recipients for this campaign: ' + existing.length);
        enrollRecipients(c);
        const afterEnroll = getRecipientsForCampaign(c.campaign_id);
        console.log('  Recipients after enroll: ' + afterEnroll.length);
        warnIfPickerStale(c);
        continue;
      }

      if (c.status === 'active') {
        console.log('  Status is active, picker_locked: ' + c.picker_locked);
        if (c.picker_locked !== 'TRUE' && c.picker_locked !== true) {
          console.log('  picker_locked is FALSE, refusing to send');
          log('skip_active_unlocked', `Campaign ${c.campaign_id} is active but picker_locked is FALSE. Refusing to send.`);
          continue;
        }
        sendInitialDms(c);
        checkExternalSignals(c);
        sendReminders(c);
        closeIfPastEndDate(c);
      }
    } catch (e) {
      console.log('  ERROR in campaign ' + c.campaign_id + ': ' + e.message);
      console.log('  Stack: ' + e.stack);
      log('orchestrator_error', `${c.campaign_id}: ${e.message}`);
    }
  }

  console.log('=== orchestrator() finished ===');
}

// ─────────────────────────────────────────────────────────────────────────────
// ENROLLMENT
// ─────────────────────────────────────────────────────────────────────────────

function enrollRecipients(campaign) {
  const eligibleEmployees = resolveAudienceFilter(campaign.audience_filter, campaign);
  const existing = getRecipientsForCampaign(campaign.campaign_id);
  const existingIds = new Set(existing.map(r => r.slack_user_id));

  let enrolledCount = 0;
  for (const emp of eligibleEmployees) {
    if (!emp.slack_user_id) {
      console.log('  Skipping empty slack_user_id row');
      continue;
    }
    if (existingIds.has(emp.slack_user_id)) {
      console.log('  Already enrolled: ' + emp.slack_user_id);
      continue;
    }
    createRecipientRow(campaign.campaign_id, emp);
    enrolledCount++;
    console.log('  Enrolled: ' + emp.name + ' (' + emp.slack_user_id + ')');
  }
  console.log('  Total newly enrolled: ' + enrolledCount);
}

function warnIfPickerStale(campaign) {
  // Post a nudge to ops channel if a campaign has been sitting in picker_open >24h
  const recipients = getRecipientsForCampaign(campaign.campaign_id);
  if (recipients.length === 0) return;

  // Find earliest queued row by enrolled_at if present
  const queued = recipients.filter(r => r.status === 'queued');
  if (queued.length === 0) return;

  // Use the first row's enrolled_at (or last_dm_at as fallback)
  const enrolledTimestamps = queued
    .map(r => r.enrolled_at)
    .filter(Boolean)
    .map(parseDate);

  if (enrolledTimestamps.length === 0) return;

  const earliest = new Date(Math.min(...enrolledTimestamps.map(d => d.getTime())));
  const hoursSinceEnroll = (new Date() - earliest) / 3600000;

  if (hoursSinceEnroll > 24 && hoursSinceEnroll < 25 && DEFAULT_FOLLOWUP_CHANNEL) {
    slackApi('chat.postMessage', {
      channel: DEFAULT_FOLLOWUP_CHANNEL,
      text: `:hourglass_flowing_sand: Campaign *${campaign.name}* has been in \`picker_open\` for ~24h with ${recipients.length} recipients queued. Review the Recipients tab and flip \`picker_locked=TRUE\` + \`status=active\` to send.`,
    });
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// INITIAL SEND
// ─────────────────────────────────────────────────────────────────────────────

function sendInitialDms(campaign) {
  const today = new Date();
  if (today < parseDate(campaign.start_date)) {
    console.log('  Start date not reached yet');
    return;
  }

  const recipients = getRecipientsForCampaign(campaign.campaign_id);

  for (const r of recipients) {
    if (r.status === 'queued' && (r.manually_excluded === 'TRUE' || r.manually_excluded === true)) {
      updateRecipientRow(r.recipient_id, { status: 'excluded' });
      continue;
    }
    if (r.status !== 'queued') continue;

    sendCampaignDm(campaign, r);
    Utilities.sleep(SEND_DELAY_MS);
  }
}

function sendCampaignDm(campaign, recipient) {
  const firstName = getSlackFirstName(recipient.slack_user_id);
  const body = (campaign.message || '')
    .replace(/{first_name}/g, firstName)
    .replace(/{campaign_name}/g, campaign.name);

  const blocks = [
    { type: 'section', text: { type: 'mrkdwn', text: body } },
  ];

  // Render links as buttons
  if (campaign.links) {
    const pairs = campaign.links.split(',').map(s => s.trim()).filter(Boolean);
    const buttons = pairs.map((pair, i) => {
      const [url, label] = pair.split('|').map(s => s.trim());
      return {
        type: 'button',
        text: { type: 'plain_text', text: label || `Link ${i + 1}` },
        url: url,
        action_id: `link_${i}`,
      };
    });
    if (buttons.length > 0) {
      blocks.push({ type: 'actions', elements: buttons.slice(0, 5) });
    }
  }

  // Reaction guide
  const reactGuide = buildReactGuide(campaign);
  if (reactGuide) {
    blocks.push({
      type: 'context',
      elements: [{ type: 'mrkdwn', text: reactGuide }],
    });
  }

  const result = slackApi('chat.postMessage', {
    channel: recipient.slack_user_id,
    text: body,
    blocks: blocks,
  });

  if (!result.ok) {
    log('send_failed', `${recipient.slack_user_id}: ${result.error}`);
    return;
  }

  updateRecipientRow(recipient.recipient_id, {
    status: 'sent',
    slack_message_ts: result.ts,
    sent_at: new Date(),
    last_dm_at: new Date(),
  });

  log('dm_sent', `${recipient.name || recipient.slack_user_id} | campaign ${campaign.campaign_id} | ts ${result.ts}`);
}

function buildReactGuide(campaign) {
  const parts = [];
  const completeEmoji = extractEmoji(campaign.completion_signal);
  if (completeEmoji) {
    parts.push(`React ${completeEmoji} when done`);
  }
  if (campaign.question_emoji) {
    parts.push(`${campaign.question_emoji} if you have a question`);
  }
  if (campaign.optout_emoji) {
    parts.push(`${campaign.optout_emoji} to opt out`);
  }
  if (parts.length === 0) return null;
  return `_${parts.join(' · ')}_`;
}

// ─────────────────────────────────────────────────────────────────────────────
// SIGNAL CHECKING (Survey Bot integration)
// ─────────────────────────────────────────────────────────────────────────────

function checkExternalSignals(campaign) {
  const requirements = parseCompletionSignal(campaign.completion_signal);
  if (!requirements) return;

  const surveyChecks = requirements.terms.filter(t => t.type === 'survey_complete');
  if (surveyChecks.length === 0) return;

  // Pull Survey Bot's Audience tab once. If it doesn't exist (Survey Bot not built yet),
  // this gracefully returns empty array.
  const surveyAudience = readSurveyBotAudience();
  if (surveyAudience.length === 0) {
    console.log('  Survey Bot audience is empty (or tab missing) — skipping survey signal check');
    return;
  }

  const recipients = getRecipientsForCampaign(campaign.campaign_id);

  for (const r of recipients) {
    if (r.status !== 'sent' && r.status !== 'needs_followup') continue;

    let signalsReceived = parseSignalSet(r.signals_received);

    for (const check of surveyChecks) {
      const surveyId = check.value;
      const matchingRow = surveyAudience.find(a =>
        a.survey_id === surveyId
        && a.slack_user_id === r.slack_user_id
        && a.status === 'complete'
      );
      if (matchingRow) {
        signalsReceived.add(`survey_complete:${surveyId}`);
      }
    }

    const signalStr = Array.from(signalsReceived).join('|');
    if (signalStr !== r.signals_received) {
      updateRecipientRow(r.recipient_id, {
        signals_received: signalStr,
        last_signal_at: new Date(),
      });
      evaluateCompletion(campaign, r.recipient_id);
    }
  }
}

function evaluateCompletion(campaign, recipient_id) {
  const r = getRecipientRow(recipient_id);
  if (!r) return;
  const requirements = parseCompletionSignal(campaign.completion_signal);
  const signals = parseSignalSet(r.signals_received);

  if (evaluateAst(requirements, signals)) {
    updateRecipientRow(recipient_id, {
      status: 'acknowledged',
      acknowledged_at: new Date(),
      ack_via: inferAckVia(signals),
    });
    log('acknowledged', `${r.name || r.slack_user_id} | ${campaign.campaign_id}`);

    if (campaign.notify_on_complete) {
      slackApi('chat.postMessage', {
        channel: campaign.notify_on_complete,
        text: `:white_check_mark: ${r.name || r.slack_user_id} completed *${campaign.name}* (via ${inferAckVia(signals)})`,
      });
    }
  }
}

function inferAckVia(signals) {
  const hasEmoji = Array.from(signals).some(s => s.startsWith('emoji:'));
  const hasSurvey = Array.from(signals).some(s => s.startsWith('survey_complete:'));
  const hasManual = Array.from(signals).some(s => s.startsWith('manual:'));
  if (hasEmoji && hasSurvey) return 'both';
  if (hasEmoji) return 'emoji';
  if (hasSurvey) return 'survey';
  if (hasManual) return 'manual';
  return 'unknown';
}

// ─────────────────────────────────────────────────────────────────────────────
// COMPLETION SIGNAL DSL PARSER
// ─────────────────────────────────────────────────────────────────────────────

function parseCompletionSignal(str) {
  if (!str) return null;
  const trimmed = String(str).trim();
  const orParts = trimmed.split(/\s+OR\s+/i);
  if (orParts.length > 1) {
    return { op: 'OR', terms: orParts.flatMap(p => parseAndChain(p)) };
  }
  return { op: 'AND', terms: parseAndChain(trimmed) };
}

function parseAndChain(str) {
  return str.split(/\s+AND\s+/i).map(t => parseTerm(t.trim()));
}

function parseTerm(t) {
  const colonIdx = t.indexOf(':');
  if (colonIdx === -1) return { type: t, value: '' };
  return { type: t.substring(0, colonIdx).trim(), value: t.substring(colonIdx + 1).trim() };
}

function evaluateAst(ast, signalSet) {
  if (!ast) return false;
  if (ast.op === 'AND') {
    return ast.terms.every(term => signalMatch(term, signalSet));
  } else if (ast.op === 'OR') {
    return ast.terms.some(term => signalMatch(term, signalSet));
  }
  return false;
}

function signalMatch(term, signalSet) {
  const needle = `${term.type}:${term.value}`;
  return signalSet.has(needle);
}

function parseSignalSet(str) {
  if (!str) return new Set();
  return new Set(String(str).split('|').map(s => s.trim()).filter(Boolean));
}

function extractEmoji(completionSignal) {
  const ast = parseCompletionSignal(completionSignal);
  if (!ast) return null;
  const emojiTerm = ast.terms.find(t => t.type === 'emoji');
  return emojiTerm ? emojiTerm.value : null;
}

// ─────────────────────────────────────────────────────────────────────────────
// REMINDERS (personalized to what each person is missing)
// ─────────────────────────────────────────────────────────────────────────────

function sendReminders(campaign) {
  const reminderDays = String(campaign.reminder_days || '').split(',')
    .map(s => parseFloat(s.trim())).filter(n => !isNaN(n));
  if (reminderDays.length === 0) return;

  const recipients = getRecipientsForCampaign(campaign.campaign_id);
  const now = new Date();

  for (const r of recipients) {
    if (r.status !== 'sent') continue; // Don't bug people who flagged a question
    if (!r.last_dm_at) continue;

    const daysSinceLast = (now - parseDate(r.last_dm_at)) / 86400000;
    const sent = parseInt(r.reminders_sent || 0);
    if (sent >= reminderDays.length) continue;
    if (daysSinceLast < reminderDays[sent]) continue;

    sendPersonalizedReminder(campaign, r, sent + 1, sent + 1 === reminderDays.length);
  }
}

function sendPersonalizedReminder(campaign, recipient, reminderNum, isFinal) {
  const firstName = getSlackFirstName(recipient.slack_user_id);
  const signals = parseSignalSet(recipient.signals_received);
  const requirements = parseCompletionSignal(campaign.completion_signal);

  const missing = requirements.terms.filter(t => !signalMatch(t, signals));
  const missingDescriptions = missing.map(describeTerm);

  let text;
  if (missingDescriptions.length === requirements.terms.length) {
    text = `Hey ${firstName}, just floating this up — *${campaign.name}* is still open. When you have a sec, please ${missingDescriptions.join(' AND ')}.`;
  } else {
    text = `Hey ${firstName}, you're almost there on *${campaign.name}* — just need: ${missingDescriptions.join(' AND ')}. Quick and easy when you have a moment.`;
  }

  if (isFinal) {
    const endDate = formatDate(parseDate(campaign.end_date));
    text += ` This closes ${endDate} — last nudge from me. If something's blocking you, just reply and let me know.`;
  }

  slackApi('chat.postMessage', {
    channel: recipient.slack_user_id,
    text: text,
  });

  updateRecipientRow(recipient.recipient_id, {
    last_dm_at: new Date(),
    reminders_sent: reminderNum,
  });

  log('reminder', `R${reminderNum} | ${recipient.name || recipient.slack_user_id} | ${campaign.campaign_id} | missing: ${missingDescriptions.join(', ')}`);

  if (isFinal && campaign.followup_channel) {
    slackApi('chat.postMessage', {
      channel: campaign.followup_channel,
      text: `:wave: *${firstName}* (${recipient.name || recipient.slack_user_id}) hasn't completed *${campaign.name}* after ${reminderNum} reminders. Still missing: ${missingDescriptions.join(', ')}.`,
    });
  }
}

function describeTerm(term) {
  if (term.type === 'emoji') return `react ${term.value} on the original DM`;
  if (term.type === 'survey_complete') return `complete the *${term.value}* survey`;
  if (term.type === 'manual') return 'follow up with Nick';
  return `complete ${term.type}`;
}

// ─────────────────────────────────────────────────────────────────────────────
// CAMPAIGN CLOSE
// ─────────────────────────────────────────────────────────────────────────────

function closeIfPastEndDate(campaign) {
  const today = new Date();
  if (!campaign.end_date) return;
  if (today < parseDate(campaign.end_date)) return;

  const recipients = getRecipientsForCampaign(campaign.campaign_id);

  for (const r of recipients) {
    if (['sent', 'needs_followup'].includes(r.status)) {
      updateRecipientRow(r.recipient_id, { status: 'expired' });
    }
  }

  updateCampaignStatus(campaign.campaign_id, 'complete');
  postCampaignSummary(campaign, recipients);
}

function postCampaignSummary(campaign, recipients) {
  const total = recipients.length;
  const ack = recipients.filter(r => r.status === 'acknowledged').length;
  const followup = recipients.filter(r => r.status === 'needs_followup').length;
  const optedOut = recipients.filter(r => r.status === 'opted_out').length;
  const excluded = recipients.filter(r => r.status === 'excluded').length;
  const expired = recipients.filter(r => r.status === 'expired').length;
  const denominator = total - excluded - optedOut;
  const ackPct = denominator > 0 ? Math.round((ack / denominator) * 100) : 0;

  const text = `:bar_chart: *${campaign.name}* closed.\n` +
    `Acknowledged: *${ack}/${denominator}* (${ackPct}%)\n` +
    `Needs follow-up: ${followup} · Opted out: ${optedOut} · Excluded: ${excluded} · Expired: ${expired}\n\n` +
    `Open the Recipients tab and filter on \`campaign_id=${campaign.campaign_id}\` for per-person detail.`;

  const channel = campaign.summary_channel || campaign.followup_channel || DEFAULT_FOLLOWUP_CHANNEL;
  if (channel) {
    slackApi('chat.postMessage', { channel: channel, text: text });
  }
  log('campaign_closed', `${campaign.campaign_id}: ${ack}/${denominator} acknowledged`);
}

// ─────────────────────────────────────────────────────────────────────────────
// SLACK EVENTS WEBHOOK
// ─────────────────────────────────────────────────────────────────────────────

function doPost(e) {
  const body = JSON.parse(e.postData.contents);

  // Slack URL verification handshake
  if (body.type === 'url_verification') {
    return ContentService.createTextOutput(body.challenge);
  }

  if (body.type === 'event_callback' && body.event) {
    try {
      handleSlackEvent(body.event);
    } catch (err) {
      log('event_error', err.message);
    }
  }

  return ContentService.createTextOutput('OK');
}

function handleSlackEvent(event) {
  if (event.type === 'reaction_added') {
    handleReactionAdded(event);
  } else if (event.type === 'reaction_removed') {
    handleReactionRemoved(event);
  }
}

function handleReactionAdded(event) {
  const messageTs = event.item.ts;
  const slackUserId = event.user;
  const emojiName = event.reaction;

  const recipient = findRecipientByMessageTs(messageTs);
  if (!recipient) {
    log('react_orphan', `Reaction on ts ${messageTs} doesn't match any recipient`);
    return;
  }
  if (recipient.slack_user_id !== slackUserId) {
    return;
  }

  const campaign = getCampaign(recipient.campaign_id);
  if (!campaign) return;

  const emojiChar = nameToEmoji(emojiName);

  // Opt out
  if (campaign.optout_emoji && (emojiChar === campaign.optout_emoji || emojiName === emojiToName(campaign.optout_emoji))) {
    updateRecipientRow(recipient.recipient_id, { status: 'opted_out', last_signal_at: new Date() });
    slackApi('chat.postMessage', {
      channel: recipient.slack_user_id,
      text: 'Got it — opted out of this one. Ping me if anything changes.',
    });
    log('opted_out', `${recipient.name || recipient.slack_user_id} | ${recipient.campaign_id}`);
    return;
  }

  // Needs follow-up
  if (campaign.question_emoji && (emojiChar === campaign.question_emoji || emojiName === emojiToName(campaign.question_emoji))) {
    updateRecipientRow(recipient.recipient_id, { status: 'needs_followup', last_signal_at: new Date() });
    const channel = campaign.followup_channel || DEFAULT_FOLLOWUP_CHANNEL;
    if (channel) {
      slackApi('chat.postMessage', {
        channel: channel,
        text: `:raising_hand: *${recipient.name || recipient.slack_user_id}* has a question about *${campaign.name}*. DM them directly.`,
      });
    }
    log('needs_followup', `${recipient.name || recipient.slack_user_id} | ${recipient.campaign_id}`);
    return;
  }

  // Completion emoji
  const completionEmoji = extractEmoji(campaign.completion_signal);
  if (completionEmoji && (emojiChar === completionEmoji || emojiName === emojiToName(completionEmoji))) {
    let signals = parseSignalSet(recipient.signals_received);
    signals.add(`emoji:${completionEmoji}`);
    updateRecipientRow(recipient.recipient_id, {
      signals_received: Array.from(signals).join('|'),
      last_signal_at: new Date(),
    });
    evaluateCompletion(campaign, recipient.recipient_id);
  } else {
    log('react_unmapped', `${recipient.name || recipient.slack_user_id} reacted with :${emojiName}: which isn't configured for ${recipient.campaign_id}`);
  }
}

function handleReactionRemoved(event) {
  const recipient = findRecipientByMessageTs(event.item.ts);
  if (!recipient) return;
  if (recipient.slack_user_id !== event.user) return;
  const campaign = getCampaign(recipient.campaign_id);
  if (!campaign) return;

  const emojiChar = nameToEmoji(event.reaction);
  let signals = parseSignalSet(recipient.signals_received);
  signals.delete(`emoji:${emojiChar}`);
  updateRecipientRow(recipient.recipient_id, {
    signals_received: Array.from(signals).join('|'),
  });
}

const EMOJI_MAP = {
  'white_check_mark': '✅',
  'heavy_check_mark': '✔',
  'raising_hand': '🙋',
  'no_entry_sign': '🚫',
  'thumbsup': '👍',
  '+1': '👍',
  'eyes': '👀',
  'wave': '👋',
};

function nameToEmoji(name) {
  return EMOJI_MAP[name] || name;
}

function emojiToName(emoji) {
  for (const [name, char] of Object.entries(EMOJI_MAP)) {
    if (char === emoji) return name;
  }
  return null;
}

// ─────────────────────────────────────────────────────────────────────────────
// AUDIENCE FILTERS (slim roster: name, slack_user_id, leadership)
// ─────────────────────────────────────────────────────────────────────────────

function resolveAudienceFilter(filter, campaign) {
  const hris = getHrisRoster();

  if (!filter || filter === 'all_employees') return hris;

  // Compound filters (AND-joined) — handle BEFORE individual filters
  if (filter.includes(' AND ')) {
    const parts = filter.split(/\s+AND\s+/);
    let result = hris;
    for (const part of parts) {
      const sub = resolveAudienceFilter(part.trim(), campaign);
      const subIds = new Set(sub.map(e => e.slack_user_id));
      result = result.filter(e => subIds.has(e.slack_user_id));
    }
    return result;
  }

  if (filter === 'leadership:true') {
    return hris.filter(e => String(e.leadership || '').toUpperCase() === 'TRUE');
  }

  if (filter.startsWith('manual:')) {
    const ids = filter.substring(7).split(',').map(s => s.trim());
    return hris.filter(e => ids.includes(e.slack_user_id));
  }

  if (filter.startsWith('slack_channel:')) {
    const channelId = filter.substring(14).trim();
    const result = slackApi('conversations.members', { channel: channelId });
    const memberIds = result.members || [];
    return hris.filter(e => memberIds.includes(e.slack_user_id));
  }

  log('filter_error', `Unknown filter: ${filter}`);
  return [];
}

// ─────────────────────────────────────────────────────────────────────────────
// SHEET HELPERS
// ─────────────────────────────────────────────────────────────────────────────

function getSheet(name) {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
}

function readTab(name) {
  const sheet = getSheet(name);
  if (!sheet) {
    console.log('Tab not found: ' + name);
    return [];
  }
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function getCampaign(campaign_id) {
  return readTab(TABS.CAMPAIGNS).find(c => c.campaign_id === campaign_id);
}

function updateCampaignStatus(campaign_id, newStatus) {
  const sheet = getSheet(TABS.CAMPAIGNS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('campaign_id');
  const statusIdx = headers.indexOf('status');
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIdx] === campaign_id) {
      sheet.getRange(i + 1, statusIdx + 1).setValue(newStatus);
      return;
    }
  }
}

function getRecipientsForCampaign(campaign_id) {
  return readTab(TABS.RECIPIENTS).filter(r => r.campaign_id === campaign_id);
}

function getRecipientRow(recipient_id) {
  return readTab(TABS.RECIPIENTS).find(r => r.recipient_id === recipient_id);
}

function findRecipientByMessageTs(ts) {
  return readTab(TABS.RECIPIENTS).find(r => r.slack_message_ts === ts);
}

function createRecipientRow(campaign_id, employee) {
  const sheet = getSheet(TABS.RECIPIENTS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Build row by header name so it works regardless of column order
  const newRow = new Array(headers.length).fill('');
  const set = (col, val) => {
    const idx = headers.indexOf(col);
    if (idx !== -1) newRow[idx] = val;
  };

  const id = Utilities.getUuid();
  set('recipient_id', id);
  set('campaign_id', campaign_id);
  set('slack_user_id', employee.slack_user_id);
  set('name', employee.name || employee.slack_user_id);
  set('manually_excluded', 'FALSE');
  set('status', 'queued');
  set('reminders_sent', 0);
  set('enrolled_at', new Date());

  sheet.appendRow(newRow);
  return id;
}

function updateRecipientRow(recipient_id, updates) {
  const sheet = getSheet(TABS.RECIPIENTS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('recipient_id');
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIdx] === recipient_id) {
      for (const key in updates) {
        const colIdx = headers.indexOf(key);
        if (colIdx !== -1) {
          sheet.getRange(i + 1, colIdx + 1).setValue(updates[key]);
        }
      }
      return;
    }
  }
}

function readSurveyBotAudience() {
  try {
    const ss = SpreadsheetApp.openById(SURVEY_BOT_SHEET_ID);
    const sheet = ss.getSheetByName('Audience');
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    const headers = data[0];
    return data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
  } catch (e) {
    console.log('readSurveyBotAudience error: ' + e.message);
    return [];
  }
}

function getHrisRoster() {
  const ss = SpreadsheetApp.openById(HRIS_SHEET_ID);
  const sheet = ss.getSheetByName(TABS.ROSTER);
  if (!sheet) {
    log('roster_error', 'HRIS_Roster tab not found');
    return [];
  }
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  }).filter(e => e.slack_user_id); // only require slack_user_id
}

function getConfigValue(key) {
  const sheet = getSheet(TABS.CONFIG);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      return String(data[i][1]);
    }
  }
  return null;
}

function log(event, message) {
  const sheet = getSheet(TABS.LOG);
  if (sheet) {
    sheet.appendRow([new Date(), event, message]);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// SLACK API
// ─────────────────────────────────────────────────────────────────────────────

function slackApi(method, payload) {
  const url = `https://slack.com/api/${method}`;
  const options = {
    method: 'post',
    contentType: 'application/json; charset=utf-8',
    headers: { Authorization: `Bearer ${SLACK_TOKEN}` },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };
  const response = UrlFetchApp.fetch(url, options);
  const result = JSON.parse(response.getContentText());
  if (!result.ok) {
    log('slack_api_error', `${method}: ${result.error}`);
  }
  return result;
}

function getSlackFirstName(slackUserId) {
  const result = slackApi('users.info', { user: slackUserId });
  if (!result.ok) return 'there';
  return result.user.profile.first_name || result.user.profile.real_name || 'there';
}

// ─────────────────────────────────────────────────────────────────────────────
// UTIL
// ─────────────────────────────────────────────────────────────────────────────

function parseDate(v) {
  if (v instanceof Date) return v;
  return new Date(v);
}

function formatDate(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMM d');
}
