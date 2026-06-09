/**
 * Danny Bank Automation - Google Apps Script
 * RECOVERY RELEASE (v5.4) - Grounded Evidence Responses and Readable Analytics
 */

const GEMINI_SETTING_KEY = 'GEMINI_API_KEY';
const GEMINI_KEY_MIGRATED_MARKER = 'Stored securely in Script Properties';
const LEGACY_GEMINI_KEY_MIGRATED_MARKER = '[Stored in Script Properties]';
const LAST_AI_USAGE_KEY = 'LAST_AI_USAGE';
const CHAT_CONTEXT_KEY = 'CHAT_CONTEXT_HISTORY';
const WEEKDAY_ORDER = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
const MAX_TOP_ITEMS = 8;
const MAX_CATEGORY_EXAMPLES = 3;
const MATRIX_TOP_CATEGORY_COUNT = 5;
const MATRIX_TOP_ACCOUNT_COUNT = 5;
const MAX_VISIBLE_TABLE_ROWS = 8;
const MAX_EVIDENCE_TRANSACTIONS = 14;
const MAX_CONTEXT_TURNS = 4;
const MAX_FULL_BREAKDOWN_ITEMS = 24;
const MAX_GROUPED_CATEGORY_EXAMPLES = 4;
const MAX_GROUNDED_TABLE_ROWS = 18;
const FULL_LEDGER_TRANSACTION_THRESHOLD = 360;
const MAX_LEDGER_CONTEXT_TRANSACTIONS = 360;
const MAX_TOOL_TRANSACTION_RESULTS = 200;
const GEMINI_MODEL_CHAIN = ['gemini-2.5-flash', 'gemini-2.5-flash-lite'];
const ANALYTICS_LAYOUT = {
  overview: { row: 1, col: 1 },
  monthlySummary: { row: 1, col: 5 },
  weekdaySummary: { row: 1, col: 13 },
  weeklySummary: { row: 1, col: 17 },
  weekdayChart: { row: 1, col: 22 },
  monthlyCashflowChart: { row: 1, col: 26 },
  topCategories: { row: 20, col: 1 },
  topAccounts: { row: 20, col: 5 },
  topMerchants: { row: 20, col: 9 },
  weekendSummary: { row: 20, col: 13 },
  topCategoriesChart: { row: 20, col: 17 },
  topAccountsChart: { row: 20, col: 21 },
  topMerchantsChart: { row: 20, col: 25 },
  monthlyCategoryMatrix: { row: 45, col: 1 },
  monthlyAccountMatrix: { row: 45, col: 8 },
  weekendMonthlyCompare: { row: 45, col: 15 },
  categoryDriftChart: { row: 45, col: 20 },
  weekendMonthlyCompareChart: { row: 45, col: 24 },
  anomalies: { row: 80, col: 1 },
  recurring: { row: 80, col: 8 },
  categoryDrift: { row: 80, col: 14 }
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🏦 Bank Automation')
    .addItem('📊 Open Command Center', 'showSidebar')
    .addSeparator()
    .addItem('📈 Refresh Dashboard & Visuals', 'refreshVisuals')
    .addItem('⚙️ Initial Setup / Repair', 'initialSetup')
    .addItem('🔑 Reset Gemini Key Storage', 'resetGeminiKeyStorage')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Danny Bank Intelligence')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

function initialSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = [
    { name: 'Transactions', headers: ['Transaction ID', 'Date', 'Name', 'Amount', 'Category', 'Account', 'Pending'], color: '#0d1117', hidden: false },
    { name: 'AI Insights Log', headers: ['Date', 'Original Insight', 'Summary'], color: '#8957e5', hidden: false },
    { name: 'Settings', headers: ['Setting', 'Value'], color: '#30363d', hidden: false },
    { name: 'Rules', headers: ['Rule ID', 'Enabled', 'Rule Type', 'Match Type', 'Match Value', 'Treatment', 'Notes'], color: '#0f766e', hidden: false },
    { name: 'Dashboard', headers: ['Dashboard'], color: '#0d1117', hidden: false },
    { name: 'Insights', headers: ['Insights'], color: '#1f2937', hidden: false },
    { name: 'Analytics', headers: ['Section', 'Value'], color: '#161b22', hidden: true }
  ];

  sheets.forEach(function(sheetConfig) {
    const sheet = ensureSheet_(ss, sheetConfig.name);
    sheet.getRange(1, 1, 1, sheetConfig.headers.length)
      .setValues([sheetConfig.headers])
      .setBackground(sheetConfig.color)
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
    if (sheetConfig.hidden) {
      sheet.hideSheet();
    } else {
      sheet.showSheet();
    }
  });

  const settings = ss.getSheetByName('Settings');
  if (!getSetting_(GEMINI_SETTING_KEY)) {
    setSetting_(GEMINI_SETTING_KEY, '');
  }

  setupRulesSheet_(ss.getSheetByName('Rules'));
  setupDashboard_(ss.getSheetByName('Dashboard'));
  setupInsights_(ss.getSheetByName('Insights'));
  SpreadsheetApp.getUi().alert('Intelligence Engine v5.4 Ready!');
}

function refreshVisuals() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ensureSheet_(ss, 'Dashboard');
  const insights = ensureSheet_(ss, 'Insights');
  const analytics = ensureSheet_(ss, 'Analytics');
  const records = getTransactionRecords_();

  if (records.length === 0) {
    return 'No transaction data found.';
  }

  const model = buildAnalyticsModel_(records);
  const sections = buildAnalyticsSections_(model);
  writeAnalyticsSections_(analytics, sections);
  renderDashboard_(dashboard, model, sections);
  renderInsights_(insights, model, sections);
  ensureGeminiKeyStatus_();
  clearChatHistory_();
  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert('Command Center v5.4 Active!');
  return 'Dashboard and insights refreshed.';
}

function chatWithData(query) {
  const records = getTransactionRecords_();
  if (records.length === 0) {
    return 'No transactions are available yet. Run a sync first.';
  }

  const apiKey = getGeminiApiKey_();
  if (!apiKey) {
    return "Error: Paste your Gemini API key into Settings!B2 once. It will be stored securely after first use.";
  }

  const model = buildAnalyticsModel_(records);
  const conversationTurns = getConversationTurns_();
  const conversationQuery = buildConversationQuery_(conversationTurns, query);
  const filters = extractRetrievalFilters_(model, conversationQuery);
  const intent = parseAiIntent_(conversationQuery, filters);
  const groundedPacket = intent.needsGroundedEvidence
    ? buildGroundedEvidencePacket_(query, model, records, intent, filters)
    : null;
  try {
    if (groundedPacket) {
      const groundedResult = runGroundedGeminiSynthesis_(query, model, apiKey, conversationTurns, intent, groundedPacket);
      const finalBody = groundedPacket.verifiedText +
        (groundedResult.text ? '\n\n## ' + groundedResult.sectionTitle + '\n' + groundedResult.text : '');
      const detailParts = [
        'Verified tools: ' + groundedPacket.toolsUsed.join(', ')
      ];
      if (groundedResult.modelsUsed.length) {
        detailParts.push('Gemini models: ' + groundedResult.modelsUsed.join(' -> '));
        detailParts.push('Gemini used for narrative only');
      } else {
        detailParts.push('Gemini not used');
      }
      const mode = groundedResult.modelsUsed.length
        ? 'Gemini Synthesis with Verified Data'
        : 'Verified Data Only';
      const message = formatChatModeReply_(mode, detailParts.join(' | '), finalBody);
      saveChatHistory_(query, message);
      saveConversationTurn_(query, finalBody, {
        mode: groundedResult.modelsUsed.length ? 'grounded-gemini' : 'grounded-verified',
        modelsUsed: groundedResult.modelsUsed,
        toolsUsed: groundedPacket.toolsUsed
      });
      return message;
    }

    const result = runGeminiFinanceAssistant_(query, model, records, apiKey, conversationTurns, intent, filters, conversationQuery);
    const detailParts = ['Gemini models: ' + result.modelsUsed.join(' -> ')];
    detailParts.push(result.toolsUsed.length
      ? 'Gemini used tools: ' + result.toolsUsed.join(', ')
      : 'Gemini used without tool calls');
    const message = formatChatModeReply_('Gemini Synthesis', detailParts.join(' | '), result.text);
    saveChatHistory_(query, message);
    saveConversationTurn_(query, result.text, {
      mode: 'gemini',
      modelsUsed: result.modelsUsed,
      toolsUsed: result.toolsUsed
    });
    return message;
  } catch (e) {
    if (isGeminiApiKeyError_(e)) {
      clearGeminiKeyStorage_();
      return 'AI Error: Gemini rejected the stored API key. I cleared the stored key. Paste your current Gemini API key into Settings!B2, then run 🏦 Bank Automation -> 📈 Refresh Dashboard & Visuals or ask again from the sidebar.';
    }

    if (groundedPacket) {
      const fallbackBody = groundedPacket.verifiedText +
        (groundedPacket.fallbackAdviceText ? '\n\n## ' + groundedPacket.fallbackSectionTitle + '\n' + groundedPacket.fallbackAdviceText : '');
      const fallbackMessage = formatChatModeReply_(
        'Verified Direct Fallback',
        'Gemini unavailable or quota-limited, verified local output used',
        fallbackBody
      );
      saveChatHistory_(query, fallbackMessage);
      saveConversationTurn_(query, fallbackBody, {
        mode: 'grounded-fallback',
        toolsUsed: groundedPacket.toolsUsed
      });
      return fallbackMessage;
    }

    const scopedRecords = hasActiveFilters_(filters) ? filterTransactionsByRetrieval_(records, filters) : [];
    const answerModel = scopedRecords.length ? buildAnalyticsModel_(scopedRecords) : model;
    const directReply = buildDirectAnswer_(answerModel, conversationQuery, intent, filters);
    if (directReply) {
      const fallbackMessage = formatChatModeReply_('Verified Direct Fallback', 'Gemini unavailable or quota-limited, verified local output used', directReply);
      saveChatHistory_(query, fallbackMessage);
      saveConversationTurn_(query, directReply, { mode: 'fallback' });
      return fallbackMessage;
    }
    return 'AI Error: ' + e.message;
  }
}

function summarizeInsight(text) {
  if (shouldUseLocalSummary_(text)) {
    return buildLocalInsightSummary_(text);
  }

  const apiKey = getGeminiApiKey_();
  if (!apiKey) {
    return buildLocalInsightSummary_(text);
  }

  try {
    return _callGemini(buildGeminiRequest_(
      'Summarize the following financial insight into 2 punchy sentences. Preserve the main recommendation.',
      trimForSummary_(text)
    ), apiKey);
  } catch (e) {
    return buildLocalInsightSummary_(text);
  }
}

function logInsightToSheet(originalInsight, summary) {
  const sheet = ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), 'AI Insights Log');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Date', 'Original Insight', 'Summary']);
  }
  sheet.appendRow([new Date(), originalInsight, summary]);
  return 'Logged! ✅';
}

function shouldUseLocalSummary_(text) {
  const value = String(text || '');
  return value.length > 9000 || value.indexOf('## Verified Data') !== -1 || value.indexOf('Mode -> Verified Data Only') !== -1;
}

function trimForSummary_(text) {
  const value = String(text || '');
  if (value.length <= 9000) {
    return value;
  }
  return value.slice(0, 8500) + '\n\n[Output truncated for summary. Full response is stored in AI Insights Log.]';
}

function buildLocalInsightSummary_(text) {
  const value = String(text || '');
  const modeMatch = value.match(/^Mode ->[^\n]+/m);
  const monthMatches = value.match(/### [A-Z][a-z]{2} \d{4} Expenses by Category/g) || [];
  const categoryMatches = value.match(/\| [^|\n]+ \| \$[\d,]+\.\d{2} \|/g) || [];
  const accountMatches = value.match(/\| [^|\n]+ ending \d{2,4} \| \$[\d,]+\.\d{2} \|/g) || [];
  const mode = modeMatch ? modeMatch[0].replace(/^Mode ->\s*/, '') : 'Verified finance response';
  const months = monthMatches.length ? monthMatches.length + ' monthly period(s)' : 'the current transaction scope';
  const accountNote = accountMatches.length ? ' Account-level totals were included with friendly card names.' : '';
  const categoryNote = categoryMatches.length ? ' Category totals and example transactions were included.' : '';
  return mode + ' logged for ' + months + '.' + accountNote + categoryNote;
}

function getChatHistory() {
  return JSON.parse(PropertiesService.getUserProperties().getProperty('chat_history') || '[]');
}

function getGeminiConfigStatus() {
  const apiKey = getGeminiApiKey_();
  if (apiKey) {
    const usage = getLastAiUsage_();
    if (usage) {
      return 'Gemini key stored securely in Script Properties. Last AI call: prompt ~' + usage.promptTokens +
        ' tok, output ' + usage.outputTokens + ' tok, total ' + usage.totalTokens + ' tok, model ' + (usage.model || 'unknown') + '.';
    }
    return 'Gemini key stored securely in Script Properties.';
  }
  return 'Gemini key missing. Paste it into Settings!B2 once to enable AI.';
}

function isGeminiApiKeyError_(error) {
  const message = String(error && error.message ? error.message : error || '').toLowerCase();
  return message.indexOf('api key not found') !== -1 ||
    message.indexOf('api_key_invalid') !== -1 ||
    message.indexOf('invalid api key') !== -1 ||
    message.indexOf('please pass a valid api key') !== -1;
}

function formatChatModeReply_(mode, detail, body) {
  return 'Mode -> ' + mode + ' (' + detail + ')\n\n' + body;
}

function buildGroundedEvidencePacket_(query, model, records, intent, filters) {
  const toolsUsed = [];
  const normalizedQuery = normalizeQueryText_(query);
  const monthArg = (filters.months && filters.months.length)
    ? filters.months.join(' ')
    : (intent.needsMonthly ? query : '');
  const singleCategory = filters.categories && filters.categories.length === 1 ? filters.categories[0] : '';
  const singleMerchant = filters.merchants && filters.merchants.length === 1 ? filters.merchants[0] : '';
  const singleAccount = filters.accounts && filters.accounts.length === 1 ? filters.accounts[0] : '';

  let monthBreakdown = null;
  if (intent.needsMonthly || (filters.months && filters.months.length)) {
    monthBreakdown = buildMonthBreakdownToolResult_(records, model, {
      month: monthArg || query,
      include_accounts: true,
      include_categories: true,
      include_examples: true
    });
    toolsUsed.push('get_month_breakdown');
  }

  let weekendAnalysis = null;
  if (intent.needsWeekend) {
    weekendAnalysis = buildWeekendAnalysisToolResult_(records, model, {
      month: monthArg || query,
      include_examples: true
    });
    toolsUsed.push('get_weekend_analysis');
  }

  let categoryBreakdowns = [];
  if (intent.needsCategoryExamples && filters.categories && filters.categories.length) {
    categoryBreakdowns = filters.categories.slice(0, 4).map(function(categoryName) {
      return {
        category: categoryName,
        result: buildCategoryBreakdownToolResult_(records, model, {
          month: monthArg || '',
          category: categoryName,
          include_examples: true
        })
      };
    });
    if (categoryBreakdowns.length) {
      toolsUsed.push('get_category_breakdown');
    }
  }

  let accountBreakdowns = [];
  if (filters.accounts && filters.accounts.length) {
    accountBreakdowns = filters.accounts.slice(0, 4).map(function(accountName) {
      return {
        account: accountName,
        result: buildAccountBreakdownToolResult_(records, model, {
          month: monthArg || '',
          account: accountName,
          include_examples: true
        })
      };
    });
    if (accountBreakdowns.length) {
      toolsUsed.push('get_account_breakdown');
    }
  }

  let transactionSearch = null;
  if (intent.needsGroupedTransactions || intent.needsTabularOutput || (intent.needsCategoryExamples && !monthBreakdown)) {
    transactionSearch = buildSearchTransactionsToolResult_(records, model, {
      month: monthArg || '',
      category: singleCategory,
      merchant: singleMerchant,
      account: singleAccount,
      expenses_only: intent.needsSpendFocus,
      limit: intent.needsTabularOutput ? MAX_GROUNDED_TABLE_ROWS : 10,
      sort: intent.needsTabularOutput ? 'recent' : 'largest'
    });
    toolsUsed.push('search_transactions');
  }

  let overview = null;
  if (intent.needsAdvice || (!monthBreakdown && !weekendAnalysis)) {
    overview = buildOverviewToolResult_(model);
    toolsUsed.push('get_overview');
  }

  const scopeRecords = hasActiveFilters_(filters) ? filterTransactionsByRetrieval_(records, filters) : records;
  const scopeModel = scopeRecords.length ? buildAnalyticsModel_(scopeRecords) : model;
  const packet = {
    query: query,
    intent: intent,
    filters: filters,
    historyModel: model,
    scopeModel: scopeModel,
    overview: overview,
    monthBreakdown: monthBreakdown,
    weekendAnalysis: weekendAnalysis,
    categoryBreakdowns: categoryBreakdowns,
    accountBreakdowns: accountBreakdowns,
    transactionSearch: transactionSearch,
    toolsUsed: toolsUsed.filter(uniqueValue_),
    needsNarrative: intent.needsAdvice || /(what stands out|where can i|how should i|tell me how|summari[sz]e.*advice|help me|analy[sz]e|analysis|insight|interpret)/.test(normalizedQuery)
  };
  packet.verifiedText = renderGroundedVerifiedResponse_(packet);
  packet.adviceContext = buildGroundedAdviceContext_(packet);
  packet.fallbackSectionTitle = intent.needsAdvice ? 'AI Advice' : 'AI Interpretation';
  packet.fallbackAdviceText = packet.needsNarrative ? renderGroundedFallbackAdvice_(packet) : '';
  return packet;
}

function runGroundedGeminiSynthesis_(query, model, apiKey, conversationTurns, intent, packet) {
  if (!packet.needsNarrative) {
    return {
      text: '',
      modelsUsed: [],
      sectionTitle: intent.needsAdvice ? 'AI Advice' : 'AI Interpretation'
    };
  }

  const request = buildGeminiRequest_(
    buildGroundedAdviceSystemPrompt_(intent),
    buildGroundedAdvicePrompt_(query, conversationTurns, packet)
  );
  const tokenEstimate = _countGeminiTokens(request, apiKey);
  const payload = _callGeminiPayload_(request, apiKey, tokenEstimate);
  const candidate = payload.candidates && payload.candidates[0];
  const parts = candidate && candidate.content && candidate.content.parts ? candidate.content.parts : [];
  const text = parts.map(function(part) {
    return part.text || '';
  }).join('').trim();
  if (!text) {
    throw new Error('Gemini returned no grounded advice text.');
  }

  return {
    text: text,
    modelsUsed: [payload._modelUsed || GEMINI_MODEL_CHAIN[0]],
    sectionTitle: intent.needsAdvice ? 'AI Advice' : 'AI Interpretation'
  };
}

function buildGroundedAdviceSystemPrompt_(intent) {
  const sectionTitle = intent.needsAdvice ? 'AI Advice' : 'AI Interpretation';
  const lines = [
    "You are Michael's Senior Wealth Strategist.",
    'The Verified Data and Grouped Transactions sections are already rendered to the user from trusted local data.',
    'Do not restate, reformat, or modify those factual sections.',
    'Write only the narrative/advice body that should appear under a markdown heading named "' + sectionTitle + '".',
    'Never invent transactions, IDs, categories, merchants, dates, or totals.',
    'Base all recommendations strictly on the verified facts and history signals provided.',
    'Do not output a preamble or repeat the user request.'
  ];

  if (intent.needsAdvice) {
    lines.push('Use four sections named Quick Wins, Subscriptions, Behavior Patterns, and Watch List.');
    lines.push('Keep each section compact and specific.');
  } else {
    lines.push('Return 2-4 compact bullets with the most important interpretation points.');
  }
  if (intent.needsCashflow) {
    lines.push('For savings-rate or cashflow questions, clearly distinguish real external income from excluded credit-card payments/transfers.');
    lines.push('If verified external income is $0.00, say the true savings rate is not calculable from the current linked accounts instead of presenting 0% as a complete answer.');
  }

  return lines.join('\n');
}

function buildGroundedAdvicePrompt_(query, conversationTurns, packet) {
  const sectionTitle = packet.intent && packet.intent.needsAdvice ? 'AI Advice' : 'AI Interpretation';
  const lines = [];
  const recentTurns = (conversationTurns || []).slice(-2);
  if (recentTurns.length) {
    lines.push('RECENT CONVERSATION');
    recentTurns.forEach(function(turn, index) {
      lines.push('User ' + (index + 1) + ' -> ' + String(turn.user || ''));
      lines.push('Assistant ' + (index + 1) + ' -> ' + truncateLabel_(String(turn.assistant || ''), 300));
    });
    lines.push('');
  }

  lines.push('CURRENT USER REQUEST');
  lines.push(String(query || ''));
  lines.push('');
  lines.push('VERIFIED ADVICE CONTEXT');
  lines.push(packet.adviceContext);
  lines.push('');
  lines.push('Write only the ' + sectionTitle + ' section body. Do not repeat the Verified Data or Grouped Transactions sections.');
  return lines.join('\n');
}

function buildGroundedAdviceContext_(packet) {
  const lines = [];
  const scopeLabel = buildFilterLabel_(packet.filters);
  lines.push('Scope -> ' + scopeLabel);

  if (packet.monthBreakdown && packet.monthBreakdown.months.length) {
    packet.monthBreakdown.months.forEach(function(monthInfo) {
      lines.push(monthInfo.month_label + ' -> Spend ' + formatCurrency_(monthInfo.spend) + ' | Income ' + formatCurrency_(monthInfo.income) + ' | Net ' + formatCurrency_(monthInfo.net));
      lines.push('Categories -> [' + formatNamedTotalList_(monthInfo.categories, MAX_FULL_BREAKDOWN_ITEMS) + ']');
      lines.push('Accounts -> [' + formatNamedTotalList_(monthInfo.accounts, MAX_FULL_BREAKDOWN_ITEMS) + ']');
    });
  }

  if (packet.weekendAnalysis) {
    lines.push('Weekend Spend -> ' + formatCurrency_(packet.weekendAnalysis.weekend_spend));
    lines.push('Weekday Spend -> ' + formatCurrency_(packet.weekendAnalysis.weekday_spend));
    lines.push('Weekend Share -> ' + packet.weekendAnalysis.weekend_share.toFixed(2) + '%');
    lines.push('Top Weekend Categories -> [' + formatNamedTotalList_(packet.weekendAnalysis.top_weekend_categories, 6) + ']');
    lines.push('Top Weekend Merchants -> [' + formatNamedTotalList_(packet.weekendAnalysis.top_weekend_merchants, 6) + ']');
  }

  if (packet.categoryBreakdowns && packet.categoryBreakdowns.length) {
    packet.categoryBreakdowns.forEach(function(entry) {
      lines.push('Category Detail -> ' + entry.category + ' -> Total ' + formatCurrency_(entry.result.total_spend));
      lines.push('Category Accounts -> [' + formatNamedTotalList_(entry.result.accounts, 8) + ']');
    });
  }

  if (packet.accountBreakdowns && packet.accountBreakdowns.length) {
    packet.accountBreakdowns.forEach(function(entry) {
      lines.push('Account Detail -> ' + entry.account + ' -> Total ' + formatCurrency_(entry.result.total_spend));
      lines.push('Account Categories -> [' + formatNamedTotalList_(entry.result.categories, 8) + ']');
    });
  }

  if (packet.transactionSearch && packet.transactionSearch.transactions && packet.transactionSearch.transactions.length) {
    lines.push('Scoped Transactions -> ' + packet.transactionSearch.transactions.length);
    lines.push('Verified Transactions -> [' + packet.transactionSearch.transactions.slice(0, 8).map(function(item) {
      return item.date + ' ' + item.name + ' ' + formatSerializedTransactionSpend_(item) + ' (' + item.category + ' / ' + item.id + ')';
    }).join('; ') + ']');
  }

  if (packet.overview) {
    lines.push('External Cashflow -> Income ' + formatCurrency_(packet.overview.total_income) + ' | Spend ' + formatCurrency_(packet.overview.total_spend) + ' | Net ' + formatCurrency_(packet.overview.net_cashflow));
    lines.push('Savings Rate Status -> ' + packet.overview.savings_rate_status);
    lines.push('Income Coverage Note -> ' + packet.overview.income_coverage_note);
    lines.push('Review-only Flags -> ' + packet.overview.review_only_count + ' transaction(s), still included in totals');
    lines.push('Excluded By Rules -> ' + packet.overview.excluded_by_rules_count + ' transaction(s), outflow ' + formatCurrency_(packet.overview.excluded_by_rules_outflow) + ', inflow ' + formatCurrency_(packet.overview.excluded_by_rules_inflow));
    lines.push('Excluded Internal Payments/Transfers -> ' + packet.overview.excluded_internal_count + ' transaction(s), outflow ' + formatCurrency_(packet.overview.excluded_internal_outflow) + ', inflow ' + formatCurrency_(packet.overview.excluded_internal_inflow));
    lines.push('History Top Categories -> [' + formatNamedTotalList_(packet.overview.top_categories, 6) + ']');
    lines.push('History Top Merchants -> [' + formatNamedTotalList_(packet.overview.top_merchants, 6) + ']');
    lines.push('Recurring Merchants -> [' + formatMerchantContextList_(packet.overview.recurring_merchants, 6) + ']');
    lines.push('Category Drift -> [' + formatDriftContextList_(packet.overview.category_drift, 6) + ']');
  }

  return lines.join('\n');
}

function renderGroundedVerifiedResponse_(packet) {
  const verifiedLines = ['## Verified Data'];
  const groupedLines = ['## Grouped Transactions'];
  verifiedLines.push('Resolved Scope -> ' + buildFilterLabel_(packet.filters));
  let wroteSection = false;
  let wroteGroupedSection = false;

  if (packet.monthBreakdown && packet.monthBreakdown.months.length) {
    wroteSection = true;
    packet.monthBreakdown.months.forEach(function(monthInfo) {
      verifiedLines.push('');
      verifiedLines.push('### ' + monthInfo.month_label + ' Expenses by Category');
      verifiedLines.push(buildNamedTotalsMarkdownTable_('Category', monthInfo.categories));
      verifiedLines.push('');
      verifiedLines.push('### ' + monthInfo.month_label + ' Account Breakdown');
      verifiedLines.push(buildNamedTotalsMarkdownTable_('Account', monthInfo.accounts));

      if (monthInfo.example_transactions && monthInfo.example_transactions.length) {
        verifiedLines.push('');
        verifiedLines.push('### Example Transactions (' + monthInfo.month_label + ')');
        verifiedLines.push(buildTransactionMarkdownTable_(monthInfo.example_transactions));
      }

      if (monthInfo.transactions_by_category && monthInfo.transactions_by_category.length) {
        wroteGroupedSection = true;
        groupedLines.push('');
        groupedLines.push('### ' + monthInfo.month_label);
        groupedLines.push(buildGroupedTransactionsMarkdown_(monthInfo.transactions_by_category));
      }
    });
  }

  if (packet.weekendAnalysis) {
    wroteSection = true;
    verifiedLines.push('');
    verifiedLines.push('### Weekend vs Weekday Summary');
    verifiedLines.push(buildNamedTotalsMarkdownTable_('Metric', [
      { name: 'Weekend Spend', total: packet.weekendAnalysis.weekend_spend },
      { name: 'Weekday Spend', total: packet.weekendAnalysis.weekday_spend },
      { name: 'Weekend Share', total: packet.weekendAnalysis.weekend_share / 100 }
    ], { percentRows: ['Weekend Share'] }));
    verifiedLines.push('');
    verifiedLines.push('### Top Weekend Categories');
    verifiedLines.push(buildNamedTotalsMarkdownTable_('Category', packet.weekendAnalysis.top_weekend_categories));
    verifiedLines.push('');
    verifiedLines.push('### Top Weekend Merchants');
    verifiedLines.push(buildNamedTotalsMarkdownTable_('Merchant', packet.weekendAnalysis.top_weekend_merchants));
    if (packet.weekendAnalysis.example_transactions && packet.weekendAnalysis.example_transactions.length) {
      verifiedLines.push('');
      verifiedLines.push('### Example Weekend Transactions');
      verifiedLines.push(buildTransactionMarkdownTable_(packet.weekendAnalysis.example_transactions));
    }
  }

  if (packet.categoryBreakdowns && packet.categoryBreakdowns.length) {
    wroteSection = true;
    packet.categoryBreakdowns.forEach(function(entry) {
      verifiedLines.push('');
      verifiedLines.push('### Category Detail: ' + entry.category);
      verifiedLines.push(buildNamedTotalsMarkdownTable_('Metric', [
        { name: 'Total Spend', total: entry.result.total_spend }
      ]));
      if (entry.result.categories && entry.result.categories.length) {
        verifiedLines.push('');
        verifiedLines.push('#### Detailed Category Mix');
        verifiedLines.push(buildNamedTotalsMarkdownTable_('Category', entry.result.categories));
      }
      if (entry.result.accounts && entry.result.accounts.length) {
        verifiedLines.push('');
        verifiedLines.push('#### Account Breakdown');
        verifiedLines.push(buildNamedTotalsMarkdownTable_('Account', entry.result.accounts));
      }
      if (entry.result.example_transactions && entry.result.example_transactions.length) {
        verifiedLines.push('');
        verifiedLines.push('#### Example Transactions');
        verifiedLines.push(buildTransactionMarkdownTable_(entry.result.example_transactions));
      }
    });
  }

  if (packet.accountBreakdowns && packet.accountBreakdowns.length) {
    wroteSection = true;
    packet.accountBreakdowns.forEach(function(entry) {
      verifiedLines.push('');
      verifiedLines.push('### Account Detail: ' + entry.account);
      verifiedLines.push(buildNamedTotalsMarkdownTable_('Metric', [
        { name: 'Total Spend', total: entry.result.total_spend }
      ]));
      if (entry.result.categories && entry.result.categories.length) {
        verifiedLines.push('');
        verifiedLines.push('#### Category Breakdown');
        verifiedLines.push(buildNamedTotalsMarkdownTable_('Category', entry.result.categories));
      }
      if (entry.result.merchants && entry.result.merchants.length) {
        verifiedLines.push('');
        verifiedLines.push('#### Merchant Breakdown');
        verifiedLines.push(buildNamedTotalsMarkdownTable_('Merchant', entry.result.merchants));
      }
      if (entry.result.example_transactions && entry.result.example_transactions.length) {
        verifiedLines.push('');
        verifiedLines.push('#### Example Transactions');
        verifiedLines.push(buildTransactionMarkdownTable_(entry.result.example_transactions));
      }
    });
  }

  if (packet.transactionSearch && packet.intent.needsTabularOutput) {
    wroteSection = true;
    verifiedLines.push('');
    verifiedLines.push('### Verified Transactions Table');
    verifiedLines.push(buildTransactionMarkdownTable_(packet.transactionSearch.table_rows || []));
  }

  if (!packet.monthBreakdown && packet.transactionSearch && packet.intent.needsGroupedTransactions && packet.transactionSearch.transactions_by_category && packet.transactionSearch.transactions_by_category.length) {
    wroteGroupedSection = true;
    groupedLines.push('');
    groupedLines.push('### Current Scope');
    groupedLines.push(buildGroupedTransactionsMarkdown_(packet.transactionSearch.transactions_by_category));
  }

  if (packet.overview && (!wroteSection || packet.intent.needsCashflow)) {
    wroteSection = true;
    verifiedLines.push('');
    verifiedLines.push('### External Cashflow Summary');
    verifiedLines.push(buildNamedTotalsMarkdownTable_('Metric', [
      { name: 'External Income', total: packet.overview.total_income },
      { name: 'External Spend', total: packet.overview.total_spend },
      { name: 'Net Cashflow', total: packet.overview.net_cashflow },
      { name: 'Daily Average Burn', total: packet.overview.daily_average_burn },
      { name: 'Excluded Internal Payments/Transfers - Outflow', total: packet.overview.excluded_internal_outflow },
      { name: 'Excluded Internal Payments/Transfers - Inflow', total: packet.overview.excluded_internal_inflow },
      { name: 'Excluded By Rules - Outflow', total: packet.overview.excluded_by_rules_outflow },
      { name: 'Excluded By Rules - Inflow', total: packet.overview.excluded_by_rules_inflow }
    ]));
    verifiedLines.push('');
    verifiedLines.push('### Cashflow Interpretation');
    verifiedLines.push('- Savings rate status: ' + packet.overview.savings_rate_status + '.');
    verifiedLines.push('- ' + packet.overview.income_coverage_note);
    verifiedLines.push('- Excluded internal payments/transfers: ' + packet.overview.excluded_internal_count + ' transaction(s). These remain in Transactions, but are removed from cashflow so card payments do not look like income.');
    verifiedLines.push('- Review-only flags: ' + packet.overview.review_only_count + ' transaction(s). These remain included in totals.');
  }

  if (!wroteSection) {
    verifiedLines.push('');
    verifiedLines.push('No verified rows matched the current scope.');
  }

  const sections = [verifiedLines.join('\n')];
  if (wroteGroupedSection) {
    sections.push(groupedLines.join('\n'));
  }
  return sections.join('\n\n');
}

function renderGroundedFallbackAdvice_(packet) {
  if (!packet.intent.needsAdvice) {
    const observations = buildHeuristicObservations_(packet.scopeModel, packet.intent);
    if (!observations.length) {
      return '- No additional interpretation beyond the verified data.';
    }
    return observations.map(function(item) {
      return '- ' + item;
    }).join('\n');
  }

  const lines = ['Advice uses your broader spending history where relevant.'];
  appendAdviceSections_(lines, packet.historyModel);
  return lines.join('\n');
}

function buildNamedTotalsMarkdownTable_(label, items, options) {
  options = options || {};
  const percentRows = options.percentRows || [];
  const rows = ['| ' + label + ' | Value |', '| :-- | --: |'];
  const source = items && items.length ? items : [{ name: 'None', total: 0 }];
  source.forEach(function(item) {
    const isPercent = percentRows.indexOf(item.name) !== -1;
    const formatted = isPercent
      ? formatPercent_(Number(item.total || 0))
      : formatCurrency_(Number(item.total || 0));
    rows.push('| ' + item.name + ' | ' + formatted + ' |');
  });
  return rows.join('\n');
}

function buildTransactionMarkdownTable_(transactions) {
  const rows = [
    '| Date | Merchant | Spend | Category | Account | Transaction ID |',
    '| :-- | :-- | --: | :-- | :-- | :-- |'
  ];
  const source = transactions && transactions.length ? transactions : [];
  if (!source.length) {
    rows.push('| N/A | No matching transactions | $0.00 | N/A | N/A | N/A |');
    return rows.join('\n');
  }
  source.forEach(function(item) {
    rows.push('| ' + item.date + ' | ' + item.name + ' | ' + formatSerializedTransactionSpend_(item) + ' | ' + item.category + ' | ' + item.account + ' | ' + item.id + ' |');
  });
  return rows.join('\n');
}

function buildGroupedTransactionsMarkdown_(groups) {
  const lines = [];
  const source = groups && groups.length ? groups : [];
  if (!source.length) {
    return '- No grouped transactions found.';
  }
  source.forEach(function(group) {
    const transactions = (group.transactions || []).map(function(item) {
      return item.name + ' ' + formatSerializedTransactionSpend_(item) + ' on ' + item.date + ' (' + item.account + ' / ' + item.id + ')';
    }).join('; ');
    lines.push('- ' + group.category + ' -> [' + transactions + ']');
  });
  return lines.join('\n');
}

function formatSerializedTransactionSpend_(item) {
  const numeric = Number(item && item.spend != null ? item.spend : item.amount || 0);
  return formatCurrency_(Math.abs(numeric));
}

function formatNamedTotalList_(items, limit) {
  const source = (items || []).slice(0, limit || 6);
  if (!source.length) {
    return 'None';
  }
  return source.map(function(item) {
    return item.name + ' ' + formatCurrency_(Number(item.total || 0));
  }).join(', ');
}

function formatMerchantContextList_(items, limit) {
  const source = (items || []).slice(0, limit || 6);
  if (!source.length) {
    return 'None';
  }
  return source.map(function(item) {
    return item.name + ' ' + formatCurrency_(Number(item.total || 0)) + ' (' + (item.count || 0) + 'x)';
  }).join(', ');
}

function formatDriftContextList_(items, limit) {
  const source = (items || []).slice(0, limit || 6);
  if (!source.length) {
    return 'None';
  }
  return source.map(function(item) {
    return item.name + ' ' + formatCurrency_(Number(item.delta || 0));
  }).join(', ');
}

function runGeminiFinanceAssistant_(query, model, records, apiKey, conversationTurns, intent, filters, conversationQuery) {
  const backgroundContext = buildGeminiBackgroundContext_(model, records, conversationQuery, intent, filters);
  const contents = buildConversationContents_(conversationTurns, query, backgroundContext);
  const toolsUsed = [];
  const modelsUsed = [];
  const systemInstruction = buildFinanceAssistantSystemPrompt_();
  const toolDeclarations = buildFinanceToolDeclarations_();

  for (let turn = 0; turn < 4; turn++) {
    const request = buildFinanceAssistantRequest_(systemInstruction, contents, toolDeclarations);
    const tokenEstimate = _countGeminiTokens(request, apiKey);
    const payload = _callGeminiPayload_(request, apiKey, tokenEstimate);
    modelsUsed.push(payload._modelUsed || GEMINI_MODEL_CHAIN[0]);
    const candidate = payload.candidates && payload.candidates[0];
    if (!candidate || !candidate.content) {
      throw new Error('Gemini returned no candidate content.');
    }

    const parts = candidate.content.parts || [];
    const functionCalls = parts.filter(function(part) {
      return part.functionCall;
    });
    const text = parts.map(function(part) {
      return part.text || '';
    }).join('').trim();

    if (!functionCalls.length) {
      if (!text) {
        throw new Error('Gemini returned no text.');
      }
      return {
        text: text,
        toolsUsed: toolsUsed.filter(uniqueValue_),
        modelsUsed: modelsUsed.filter(uniqueValue_)
      };
    }

    contents.push({ role: 'model', parts: parts });
    contents.push({
      role: 'user',
      parts: functionCalls.map(function(part) {
        const functionCall = part.functionCall;
        const args = normalizeFunctionArgs_(functionCall.args);
        const result = executeFinanceTool_(functionCall.name, args, model, records);
        toolsUsed.push(functionCall.name);
        return {
          functionResponse: {
            name: functionCall.name,
            id: functionCall.id,
            response: { result: result }
          }
        };
      })
    });
  }

  throw new Error('Gemini exceeded the tool-calling loop limit.');
}

function buildFinanceAssistantSystemPrompt_() {
  return [
    "You are Michael's Senior Wealth Strategist.",
    'You have access to exact local finance tools, recent conversation context, verified analytics, and raw transaction evidence.',
    'Use prior turns to resolve follow-up questions such as "that month", "now", "those", or "break it down further".',
    'Use your own analysis and judgment. Do not simply restate tool JSON.',
    'Call tools when you need exact scoped totals, drill-downs, tables, grouped transactions, or more transaction rows.',
    'For month, category, account, merchant, weekend, or table-style requests, prefer tools before answering.',
    'If the verified background context already answers the question, you may answer directly without tool calls.',
    'Do not invent totals, categories, accounts, merchants, dates, or transaction IDs.',
    'If the user asks for examples, include examples from the raw ledger or tool results.',
    'If a month is ambiguous, say which resolved month(s) were used.',
    'Tool output is the source of truth when exact values are requested.',
    'For savings-rate questions, verify that external income exists. Do not present 0% as a true savings rate when the only positive rows are excluded card payments/transfers.',
    'When a month breakdown is requested, list every non-zero category and every active account in scope, not just the top few.',
    'If the user asks for transactions by category, separate the full grouped breakdown from the example transaction list.',
    'Normalize raw Plaid category names into cleaner user-facing labels when possible.',
    'When advice is requested, use four sections named Quick Wins, Subscriptions, Behavior Patterns, and Watch List.',
    'When useful, provide a compact table in plain text markdown.',
    'Be concise but complete.'
  ].join('\n');
}

function buildFinanceAssistantRequest_(systemInstruction, contents, toolDeclarations) {
  return {
    systemInstruction: {
      role: 'system',
      parts: [{ text: systemInstruction }]
    },
    contents: contents,
    tools: [{ functionDeclarations: toolDeclarations }],
    toolConfig: {
      functionCallingConfig: {
        mode: 'AUTO'
      }
    },
    generationConfig: {
      temperature: 0.35,
      topP: 0.9,
      maxOutputTokens: 1800
    }
  };
}

function buildFinanceToolDeclarations_() {
  return [
    {
      name: 'get_overview',
      description: 'Get an exact financial overview including totals, savings rate, top accounts, top categories, top merchants, latest month, recurring burden, and drift.',
      parameters: { type: 'OBJECT', properties: {} }
    },
    {
      name: 'get_month_breakdown',
      description: 'Get exact month-by-month or specific-month spend, income, net, account breakdowns, category breakdowns, and example transactions. Use YYYY-MM when possible.',
      parameters: {
        type: 'OBJECT',
        properties: {
          month: { type: 'STRING' },
          include_accounts: { type: 'BOOLEAN' },
          include_categories: { type: 'BOOLEAN' },
          include_examples: { type: 'BOOLEAN' }
        }
      }
    },
    {
      name: 'get_category_breakdown',
      description: 'Get exact spend for a category, optionally scoped to a month, including example transactions.',
      parameters: {
        type: 'OBJECT',
        properties: {
          category: { type: 'STRING' },
          month: { type: 'STRING' },
          include_examples: { type: 'BOOLEAN' }
        },
        required: ['category']
      }
    },
    {
      name: 'get_account_breakdown',
      description: 'Get exact spend for an account, optionally scoped to a month, including example transactions.',
      parameters: {
        type: 'OBJECT',
        properties: {
          account: { type: 'STRING' },
          month: { type: 'STRING' },
          include_examples: { type: 'BOOLEAN' }
        },
        required: ['account']
      }
    },
    {
      name: 'get_weekend_analysis',
      description: 'Get weekend versus weekday spend, optional month scope, top weekend categories and merchants, and examples.',
      parameters: {
        type: 'OBJECT',
        properties: {
          month: { type: 'STRING' },
          include_examples: { type: 'BOOLEAN' }
        }
      }
    },
    {
      name: 'search_transactions',
      description: 'Return exact raw transactions matching optional month, category, merchant, and account filters. Use this when you need broader evidence, exhaustive examples, or a detailed table.',
      parameters: {
        type: 'OBJECT',
        properties: {
          month: { type: 'STRING' },
          category: { type: 'STRING' },
          merchant: { type: 'STRING' },
          account: { type: 'STRING' },
          expenses_only: { type: 'BOOLEAN' },
          limit: { type: 'NUMBER' },
          sort: { type: 'STRING' }
        }
      }
    },
    {
      name: 'get_recurring_merchants',
      description: 'Return recurring or subscription-like merchants based on repeated spend patterns.',
      parameters: {
        type: 'OBJECT',
        properties: {
          limit: { type: 'NUMBER' }
        }
      }
    },
    {
      name: 'get_anomalies',
      description: 'Return largest or unusual transactions worth reviewing.',
      parameters: {
        type: 'OBJECT',
        properties: {
          limit: { type: 'NUMBER' }
        }
      }
    },
    {
      name: 'get_category_drift',
      description: 'Return category changes between the latest month and the prior month.',
      parameters: {
        type: 'OBJECT',
        properties: {
          limit: { type: 'NUMBER' }
        }
      }
    }
  ];
}

function executeFinanceTool_(name, args, model, records) {
  switch (name) {
    case 'get_overview':
      return buildOverviewToolResult_(model);
    case 'get_month_breakdown':
      return buildMonthBreakdownToolResult_(records, model, args);
    case 'get_category_breakdown':
      return buildCategoryBreakdownToolResult_(records, model, args);
    case 'get_account_breakdown':
      return buildAccountBreakdownToolResult_(records, model, args);
    case 'get_weekend_analysis':
      return buildWeekendAnalysisToolResult_(records, model, args);
    case 'search_transactions':
      return buildSearchTransactionsToolResult_(records, model, args);
    case 'get_recurring_merchants':
      return {
        recurring_merchants: serializeMerchantItems_(model.recurringCandidates, normalizeLimit_(args.limit, 8))
      };
    case 'get_anomalies':
      return {
        anomalies: model.anomalies.slice(0, normalizeLimit_(args.limit, 8))
      };
    case 'get_category_drift':
      return {
        category_drift: model.categoryDrift.slice(0, normalizeLimit_(args.limit, 8))
      };
    default:
      return { error: 'Unknown tool: ' + name };
  }
}

function saveChatHistory_(userMessage, botReply) {
  const props = PropertiesService.getUserProperties();
  const history = JSON.parse(props.getProperty('chat_history') || '[]');
  history.push({ user: userMessage, bot: botReply });
  if (history.length > 8) {
    history.shift();
  }
  props.setProperty('chat_history', JSON.stringify(history));
}

function getConversationTurns_() {
  try {
    return JSON.parse(PropertiesService.getUserProperties().getProperty(CHAT_CONTEXT_KEY) || '[]');
  } catch (e) {
    return [];
  }
}

function saveConversationTurn_(userMessage, assistantMessage, meta) {
  const props = PropertiesService.getUserProperties();
  const history = getConversationTurns_();
  history.push({
    user: String(userMessage || ''),
    assistant: String(assistantMessage || ''),
    mode: meta && meta.mode ? meta.mode : '',
    modelsUsed: meta && meta.modelsUsed ? meta.modelsUsed : [],
    toolsUsed: meta && meta.toolsUsed ? meta.toolsUsed : []
  });
  while (history.length > MAX_CONTEXT_TURNS) {
    history.shift();
  }
  props.setProperty(CHAT_CONTEXT_KEY, JSON.stringify(history));
}

function clearChatHistory_() {
  const props = PropertiesService.getUserProperties();
  props.deleteProperty('chat_history');
  props.deleteProperty(CHAT_CONTEXT_KEY);
}

function setupDashboard_(sheet) {
  sheet.clear();
  sheet.setTabColor('#34c759');
  sheet.setHiddenGridlines(true);
  for (let col = 1; col <= 14; col++) {
    sheet.setColumnWidth(col, col === 1 ? 180 : 120);
  }
  sheet.getRange('A1:N1').merge()
    .setValue('Wealth Intelligence Dashboard')
    .setFontSize(20)
    .setFontWeight('bold')
    .setBackground('#0d1117')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  sheet.getRange('A2:N2').merge()
    .setValue('Executive view: external cashflow, categories, cadence, and merchant concentration (internal payments/transfers and active Rules excluded)')
    .setBackground('#111827')
    .setFontColor('#93c5fd')
    .setHorizontalAlignment('center');
  sheet.getRange('A3:H3')
    .setValues([['EXTERNAL INCOME', 'EXTERNAL SPEND', 'NET CASHFLOW', 'SAVINGS RATE', 'DAILY BURN', 'TOP MERCHANT', 'WEEKEND SHARE', 'RECURRING']])
    .setBackground('#1f2937')
    .setFontColor('#93c5fd')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.getRange('A3:H4').setBorder(true, true, true, true, true, true, '#374151', SpreadsheetApp.BorderStyle.SOLID);
}

function setupInsights_(sheet) {
  sheet.clear();
  sheet.setTabColor('#7c3aed');
  sheet.setHiddenGridlines(true);
  for (let col = 1; col <= 18; col++) {
    sheet.setColumnWidth(col, col === 1 ? 180 : 120);
  }
  sheet.getRange('A1:P1').merge()
    .setValue('Wealth Intelligence Insights')
    .setFontSize(20)
    .setFontWeight('bold')
    .setBackground('#111827')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  sheet.getRange('A2:P2').merge()
    .setValue('Deeper cuts: composition, accounts, anomalies, recurring spend, and drift')
    .setBackground('#1f2937')
    .setFontColor('#c4b5fd')
    .setHorizontalAlignment('center');
}

function renderDashboard_(sheet, model, sections) {
  setupDashboard_(sheet);
  clearCharts_(sheet);
  writeDashboardKpis_(sheet, model);
  const analytics = sheet.getParent().getSheetByName('Analytics');
  const hasIncomeSeries = monthlyCashflowHasIncome_(model);

  const charts = [
    sheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(getSectionRange_(analytics, sections.topCategoriesChart))
      .setNumHeaders(1)
      .setPosition(8, 1, 0, 0)
      .setOption('title', 'Top External Spend Categories')
      .setOption('legend', { position: 'none' })
      .setOption('hAxis', { title: 'Spend ($)' })
      .setOption('vAxis', { title: 'Category' })
      .setOption('colors', ['#2563eb'])
      .build(),
    sheet.newChart()
      .setChartType(Charts.ChartType.COMBO)
      .addRange(getSectionRange_(analytics, sections.monthlyCashflowChart))
      .setNumHeaders(1)
      .setPosition(8, 8, 0, 0)
      .setOption('title', hasIncomeSeries ? 'Monthly External Cashflow (Income vs Spend vs Net)' : 'Monthly External Spend vs Net Cashflow')
      .setOption('seriesType', 'bars')
      .setOption('hAxis', { title: 'Month', slantedText: true, slantedTextAngle: 35 })
      .setOption('vAxes', hasIncomeSeries ? {
        0: { title: 'Income / Spend ($)' },
        1: { title: 'Net Cashflow ($)' }
      } : {
        0: { title: 'External Spend ($)' },
        1: { title: 'Net Cashflow ($)' }
      })
      .setOption('legend', { position: 'top', textStyle: { fontSize: 10 } })
      .setOption('series', hasIncomeSeries ? {
        2: { type: 'line', color: '#111827', lineWidth: 3, pointSize: 6, targetAxisIndex: 1 }
      } : {
        1: { type: 'line', color: '#111827', lineWidth: 3, pointSize: 6, targetAxisIndex: 1 }
      })
      .setOption('colors', hasIncomeSeries ? ['#16a34a', '#dc2626', '#111827'] : ['#2563eb', '#111827'])
      .build(),
    sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(getSectionRange_(analytics, sections.weeklySummary, 1, 2))
      .setNumHeaders(1)
      .setPosition(25, 1, 0, 0)
      .setOption('title', 'Weekly External Spend by Calendar Week')
      .setOption('curveType', 'function')
      .setOption('legend', { position: 'none' })
      .setOption('hAxis', { title: 'Week Starting' })
      .setOption('vAxis', { title: 'Spend ($)' })
      .setOption('pointSize', 5)
      .setOption('colors', ['#0f766e'])
      .build(),
    sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(getSectionRange_(analytics, sections.weekdayChart))
      .setNumHeaders(1)
      .setPosition(25, 8, 0, 0)
      .setOption('title', 'Weekday External Spend Pattern')
      .setOption('legend', { position: 'none' })
      .setOption('hAxis', { title: 'Day of Week' })
      .setOption('vAxis', { title: 'Spend ($)' })
      .setOption('colors', ['#7c3aed'])
      .build(),
    sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(getSectionRange_(analytics, sections.weekendMonthlyCompareChart))
      .setNumHeaders(1)
      .setPosition(42, 1, 0, 0)
      .setOption('title', 'Weekend vs Weekday External Spend by Month')
      .setOption('hAxis', { title: 'Month', slantedText: true, slantedTextAngle: 35 })
      .setOption('vAxis', { title: 'Spend ($)' })
      .setOption('legend', { position: 'top', textStyle: { fontSize: 10 } })
      .setOption('colors', ['#64748b', '#f97316'])
      .build(),
    sheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(getSectionRange_(analytics, sections.topMerchantsChart))
      .setNumHeaders(1)
      .setPosition(42, 8, 0, 0)
      .setOption('title', 'Top External Spend Merchants')
      .setOption('legend', { position: 'none' })
      .setOption('hAxis', { title: 'Spend ($)' })
      .setOption('vAxis', { title: 'Merchant' })
      .setOption('colors', ['#0ea5e9'])
      .build()
  ];

  insertCharts_(sheet, charts);
}

function renderInsights_(sheet, model, sections) {
  setupInsights_(sheet);
  clearCharts_(sheet);

  const analytics = sheet.getParent().getSheetByName('Analytics');
  const charts = [
    sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(getSectionRange_(analytics, sections.monthlyCategoryMatrix))
      .setNumHeaders(1)
      .setPosition(4, 1, 0, 0)
      .setOption('title', 'Monthly External Spend by Category')
      .setOption('isStacked', true)
      .setOption('hAxis', { title: 'Month', slantedText: true, slantedTextAngle: 35 })
      .setOption('vAxis', { title: 'Spend ($)' })
      .setOption('legend', { position: 'top', textStyle: { fontSize: 10 } })
      .build(),
    sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(getSectionRange_(analytics, sections.monthlyAccountMatrix))
      .setNumHeaders(1)
      .setPosition(4, 9, 0, 0)
      .setOption('title', 'Monthly External Spend by Account')
      .setOption('isStacked', true)
      .setOption('hAxis', { title: 'Month', slantedText: true, slantedTextAngle: 35 })
      .setOption('vAxis', { title: 'Spend ($)' })
      .setOption('legend', { position: 'top', textStyle: { fontSize: 10 } })
      .build(),
    sheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(getSectionRange_(analytics, sections.topAccountsChart))
      .setNumHeaders(1)
      .setPosition(22, 1, 0, 0)
      .setOption('title', 'External Spend by Account')
      .setOption('legend', { position: 'none' })
      .setOption('hAxis', { title: 'Spend ($)' })
      .setOption('vAxis', { title: 'Account' })
      .setOption('colors', ['#2563eb'])
      .build(),
    sheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(getSectionRange_(analytics, sections.categoryDriftChart))
      .setNumHeaders(1)
      .setPosition(22, 9, 0, 0)
      .setOption('title', 'Category Drift (Latest vs Prior Month)')
      .setOption('hAxis', { title: 'Delta vs Prior Month ($)' })
      .setOption('vAxis', { title: 'Category' })
      .setOption('legend', { position: 'none' })
      .setOption('colors', ['#dc2626'])
      .build()
  ];

  insertCharts_(sheet, charts);
  writeVisibleSectionHeader_(sheet, 43, 1, 'Largest Transactions To Review', 6, '#f59e0b');
  writeVisibleSectionHeader_(sheet, 43, 8, 'Recurring Merchants', 5, '#8b5cf6');
  writeVisibleSectionHeader_(sheet, 43, 14, 'Category Drift Table', 4, '#ef4444');
  writeTable_(sheet, 44, 1, sections.anomalies.values.slice(0, MAX_VISIBLE_TABLE_ROWS + 1), '#f59e0b');
  writeTable_(sheet, 44, 8, sections.recurring.values.slice(0, MAX_VISIBLE_TABLE_ROWS + 1), '#8b5cf6');
  writeTable_(sheet, 44, 14, sections.categoryDrift.values.slice(0, MAX_VISIBLE_TABLE_ROWS + 1), '#ef4444');
}

function writeDashboardKpis_(sheet, model) {
  const recurringCount = model.recurringCandidates.length;
  const weekendShare = model.totalSpend > 0 ? model.weekendSpend / model.totalSpend : 0;
  const values = [[
    formatCurrency_(model.totalIncome),
    formatCurrency_(model.totalSpend),
    formatCurrency_(model.netCashflow),
    formatSavingsRateForDisplay_(model),
    formatCurrency_(model.dailyAverageBurn),
    truncateLabel_(model.topMerchantName || 'N/A', 16),
    formatPercent_(weekendShare),
    recurringCount + ' merchants'
  ]];
  sheet.getRange('A4:H4')
    .setValues(values)
    .setBackground('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.getRange('A5:H5').merge()
    .setValue('Analytics scope: outflow means money leaving accounts; inflow means money entering accounts. Rule exclusions remove rows from Dashboard/Insights/AI totals. Review-only rules only flag rows for attention and keep them included. Internal payment/transfer exclusions prevent credit-card payments from being double-counted as income.')
    .setBackground('#f8fafc')
    .setFontColor('#334155')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.setRowHeight(5, 42);
  writeRuleImpactSummary_(sheet, model);
}

function writeRuleImpactSummary_(sheet, model) {
  const internalCount = model.excludedTransactionCount || 0;
  const reviewCount = model.reviewOnlyTransactionCount || 0;
  const ruleCount = model.ruleExcludedTransactionCount || 0;
  const values = [
    ['RULE / EXCLUSION IMPACT', 'COUNT', 'OUTFLOW', 'INFLOW', 'DASHBOARD EFFECT', 'PLAIN ENGLISH'],
    [
      'Rules excluded',
      ruleCount,
      formatCurrency_(model.ruleExcludedCashOutflow || 0),
      formatCurrency_(model.ruleExcludedCashInflow || 0),
      'Removed from totals',
      'Enabled exclude rules remove matching rows from Dashboard, Insights, Analytics, and AI totals.'
    ],
    [
      'Review-only flags',
      reviewCount,
      'Included',
      'Included',
      'Still counted',
      'Rows match review_only rules. They stay in totals but are called out for review and AI context.'
    ],
    [
      'Internal payments/transfers',
      internalCount,
      formatCurrency_(model.excludedCashOutflow || 0),
      formatCurrency_(model.excludedCashInflow || 0),
      'Auto-excluded',
      'Credit-card payments/transfers are raw ledger rows, but excluded from cashflow to avoid fake income.'
    ]
  ];
  sheet.getRange('I3:N6')
    .setValues(values)
    .setWrap(true)
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#374151', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('I3:N3')
    .setBackground('#0f766e')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.getRange('I4:N6')
    .setBackground('#ffffff')
    .setFontColor('#111827')
    .setFontSize(9);
  sheet.getRange('J4:J6').setHorizontalAlignment('center').setFontWeight('bold');
  sheet.getRange('L4:L6').setHorizontalAlignment('center');
  if (reviewCount > 0) {
    sheet.getRange('I5:N5').setBackground('#fff7ed');
  }
  if (ruleCount > 0) {
    sheet.getRange('I4:N4').setBackground('#fef2f2');
  }
  if (internalCount > 0) {
    sheet.getRange('I6:N6').setBackground('#eff6ff');
  }
  sheet.setRowHeight(3, 24);
  sheet.setRowHeight(4, 42);
  sheet.setRowHeight(5, 46);
  sheet.setRowHeight(6, 46);
}

function buildAnalyticsModel_(records, analyticsRules) {
  const timezone = getSpreadsheetTimeZone_();
  const rules = analyticsRules || getAnalyticsRules_();
  const model = {
    timezone: timezone,
    activeRules: rules,
    transactionCount: 0,
    expenseCount: 0,
    incomeCount: 0,
    pendingCount: 0,
    excludedTransactionCount: 0,
    excludedCashOutflow: 0,
    excludedCashInflow: 0,
    ruleExcludedTransactionCount: 0,
    ruleExcludedCashOutflow: 0,
    ruleExcludedCashInflow: 0,
    reviewOnlyTransactionCount: 0,
    totalSpend: 0,
    totalIncome: 0,
    netCashflow: 0,
    dailyAverageBurn: 0,
    savingsRate: 0,
    weekendSpend: 0,
    weekdaySpendTotal: 0,
    minDate: null,
    maxDate: null,
    dayCount: 1,
    monthCount: '0.0',
    topMerchantName: 'N/A',
    weekdays: initializeWeekdayTotals_(),
    weeks: {},
    months: {},
    accounts: {},
    primaryCategories: {},
    detailedCategories: {},
    merchants: {},
    anomalies: [],
    weekendCategories: {},
    weekendMerchants: {},
    weekendExamples: [],
    recurringCandidates: [],
    categoryDrift: []
  };

  records.forEach(function(record) {
    if (!record.date) {
      return;
    }

    model.transactionCount += 1;
    if (record.pending) {
      model.pendingCount += 1;
    }

    if (!model.minDate || record.date < model.minDate) {
      model.minDate = record.date;
    }
    if (!model.maxDate || record.date > model.maxDate) {
      model.maxDate = record.date;
    }

    const monthKey = Utilities.formatDate(record.date, timezone, 'yyyy-MM');
    const weekKey = Utilities.formatDate(getWeekStart_(record.date), timezone, 'yyyy-MM-dd');
    const weekdayKey = Utilities.formatDate(record.date, timezone, 'EEEE');
    const accountLabel = formatAccountLabel_(record.account);
    const categoryLabel = formatCategoryLabel_(record.category);
    const detailedCategoryLabel = formatDetailedCategoryLabel_(record.category);
    const merchantLabel = truncateLabel_(record.name, 36);
    const monthBucket = getOrCreateMonthBucket_(model.months, monthKey);
    const weekBucket = getOrCreateWeekBucket_(model.weeks, weekKey);
    const ruleDecision = classifyRecordByRules_(record, rules, {
      account: accountLabel,
      category: categoryLabel,
      detailedCategory: detailedCategoryLabel,
      merchant: merchantLabel
    });

    if (ruleDecision.treatment === 'review_only') {
      model.reviewOnlyTransactionCount += 1;
      monthBucket.reviewOnlyCount += 1;
    }

    if (ruleDecision.treatment === 'exclude_from_analytics' || ruleDecision.treatment === 'exclude_from_cashflow') {
      model.ruleExcludedTransactionCount += 1;
      monthBucket.ruleExcludedCount += 1;
      if (record.amount > 0) {
        model.ruleExcludedCashInflow += record.amount;
        monthBucket.ruleExcludedInflow += record.amount;
      } else if (record.amount < 0) {
        const ruleMovedAmount = Math.abs(record.amount);
        model.ruleExcludedCashOutflow += ruleMovedAmount;
        monthBucket.ruleExcludedOutflow += ruleMovedAmount;
      }
      return;
    }

    const cashflowClass = classifyCashflowRecord_(record, categoryLabel, detailedCategoryLabel);

    if (cashflowClass.excludeFromCashflow) {
      model.excludedTransactionCount += 1;
      monthBucket.excludedCount += 1;
      if (record.amount > 0) {
        model.excludedCashInflow += record.amount;
        monthBucket.excludedInflow += record.amount;
      } else if (record.amount < 0) {
        const movedAmount = Math.abs(record.amount);
        model.excludedCashOutflow += movedAmount;
        monthBucket.excludedOutflow += movedAmount;
      }
      return;
    }

    if (record.amount > 0) {
      model.incomeCount += 1;
      model.totalIncome += record.amount;
      monthBucket.income += record.amount;
      monthBucket.transactionCount += 1;
      weekBucket.income += record.amount;
      weekBucket.transactionCount += 1;
      return;
    }

    if (record.amount >= 0) {
      return;
    }

    const spend = Math.abs(record.amount);
    const isWeekend = isWeekendDate_(record.date);

    model.expenseCount += 1;
    model.totalSpend += spend;
    model.weekdays[weekdayKey] = (model.weekdays[weekdayKey] || 0) + spend;
    model.accounts[accountLabel] = (model.accounts[accountLabel] || 0) + spend;
    model.primaryCategories[categoryLabel] = model.primaryCategories[categoryLabel] || createCategoryStat_(categoryLabel);
    model.primaryCategories[categoryLabel].total += spend;
    model.primaryCategories[categoryLabel].examples.push(createExample_(record, spend, timezone));
    model.detailedCategories[detailedCategoryLabel] = model.detailedCategories[detailedCategoryLabel] || createCategoryStat_(detailedCategoryLabel);
    model.detailedCategories[detailedCategoryLabel].total += spend;
    model.detailedCategories[detailedCategoryLabel].examples.push(createExample_(record, spend, timezone));
    model.merchants[merchantLabel] = model.merchants[merchantLabel] || createMerchantStat_(merchantLabel);
    model.merchants[merchantLabel].total += spend;
    model.merchants[merchantLabel].count += 1;
    model.merchants[merchantLabel].lastSeen = Utilities.formatDate(record.date, timezone, 'yyyy-MM-dd');

    monthBucket.spend += spend;
    monthBucket.transactionCount += 1;
    monthBucket.accounts[accountLabel] = (monthBucket.accounts[accountLabel] || 0) + spend;
    monthBucket.categories[categoryLabel] = (monthBucket.categories[categoryLabel] || 0) + spend;
    monthBucket.merchants[merchantLabel] = (monthBucket.merchants[merchantLabel] || 0) + spend;
    monthBucket.examples.push(createExample_(record, spend, timezone));

    weekBucket.spend += spend;
    weekBucket.transactionCount += 1;

    model.anomalies.push({
      date: Utilities.formatDate(record.date, timezone, 'yyyy-MM-dd'),
      merchant: merchantLabel,
      category: categoryLabel,
      spend: spend,
      account: accountLabel,
      id: record.id
    });

    if (isWeekend) {
      model.weekendSpend += spend;
      monthBucket.weekendSpend += spend;
      model.weekendCategories[categoryLabel] = (model.weekendCategories[categoryLabel] || 0) + spend;
      model.weekendMerchants[merchantLabel] = (model.weekendMerchants[merchantLabel] || 0) + spend;
      model.weekendExamples.push(createExample_(record, spend, timezone));
    } else {
      model.weekdaySpendTotal += spend;
      monthBucket.weekdaySpend += spend;
    }
  });

  finalizeAnalyticsModel_(model);
  return model;
}

function finalizeAnalyticsModel_(model) {
  if (model.minDate && model.maxDate) {
    model.dayCount = Math.max(1, Math.round((model.maxDate - model.minDate) / 86400000) + 1);
    model.monthCount = (model.dayCount / 30.44).toFixed(1);
    model.dailyAverageBurn = model.totalSpend / Math.max(1, model.dayCount);
  }

  model.netCashflow = model.totalIncome - model.totalSpend;
  model.savingsRate = model.totalIncome > 0 ? model.netCashflow / model.totalIncome : 0;

  model.monthKeys = Object.keys(model.months).sort();
  model.weekKeys = Object.keys(model.weeks).sort();
  model.accountList = bucketMapToSortedList_(model.accounts);
  model.categoryList = categoryMapToSortedList_(model.primaryCategories);
  model.detailedCategoryList = categoryMapToSortedList_(model.detailedCategories);
  model.merchantList = merchantMapToSortedList_(model.merchants);
  model.weekendCategoryList = bucketMapToSortedList_(model.weekendCategories);
  model.weekendMerchantList = bucketMapToSortedList_(model.weekendMerchants);
  model.topMerchantName = model.merchantList.length ? model.merchantList[0].name : 'N/A';
  model.recurringCandidates = model.merchantList.filter(function(entry) {
    return entry.count >= 2;
  }).slice(0, 10);
  model.anomalies = model.anomalies.sort(function(a, b) {
    return b.spend - a.spend;
  }).slice(0, 12);

  model.monthKeys.forEach(function(monthKey) {
    const bucket = model.months[monthKey];
    bucket.net = bucket.income - bucket.spend;
    bucket.accountList = bucketMapToSortedList_(bucket.accounts);
    bucket.categoryList = bucketMapToSortedList_(bucket.categories);
    bucket.merchantList = bucketMapToSortedList_(bucket.merchants);
    bucket.topAccount = bucket.accountList.length ? bucket.accountList[0].name : 'N/A';
    bucket.topCategory = bucket.categoryList.length ? bucket.categoryList[0].name : 'N/A';
    bucket.examples = bucket.examples.sort(function(a, b) {
      return b.amount - a.amount;
    }).slice(0, MAX_CATEGORY_EXAMPLES);
  });

  model.weekKeys.forEach(function(weekKey) {
    const bucket = model.weeks[weekKey];
    bucket.net = bucket.income - bucket.spend;
  });

  model.topCategoryNames = model.categoryList.slice(0, MATRIX_TOP_CATEGORY_COUNT).map(function(entry) {
    return entry.name;
  });
  model.topAccountNames = model.accountList.slice(0, MATRIX_TOP_ACCOUNT_COUNT).map(function(entry) {
    return entry.name;
  });
  if (!model.topCategoryNames.length) {
    model.topCategoryNames = ['Uncategorized'];
  }
  if (!model.topAccountNames.length) {
    model.topAccountNames = ['Unknown Account'];
  }

  if (model.monthKeys.length >= 2) {
    const latestKey = model.monthKeys[model.monthKeys.length - 1];
    const previousKey = model.monthKeys[model.monthKeys.length - 2];
    const latestCategories = model.months[latestKey].categories;
    const previousCategories = model.months[previousKey].categories;
    const names = Object.keys(latestCategories).concat(Object.keys(previousCategories)).filter(uniqueValue_);
    model.categoryDrift = names.map(function(name) {
      const latest = latestCategories[name] || 0;
      const previous = previousCategories[name] || 0;
      return {
        name: name,
        latest: latest,
        previous: previous,
        delta: latest - previous
      };
    }).sort(function(a, b) {
      return Math.abs(b.delta) - Math.abs(a.delta);
    }).slice(0, 8);
  } else {
    model.categoryDrift = [];
  }
}

function buildAnalyticsSections_(model) {
  return {
    overview: createSection_(ANALYTICS_LAYOUT.overview.row, ANALYTICS_LAYOUT.overview.col, buildOverviewTable_(model)),
    monthlySummary: createSection_(ANALYTICS_LAYOUT.monthlySummary.row, ANALYTICS_LAYOUT.monthlySummary.col, buildMonthlySummaryTable_(model)),
    weekdaySummary: createSection_(ANALYTICS_LAYOUT.weekdaySummary.row, ANALYTICS_LAYOUT.weekdaySummary.col, buildWeekdaySummaryTable_(model)),
    weeklySummary: createSection_(ANALYTICS_LAYOUT.weeklySummary.row, ANALYTICS_LAYOUT.weeklySummary.col, buildWeeklySummaryTable_(model)),
    weekdayChart: createSection_(ANALYTICS_LAYOUT.weekdayChart.row, ANALYTICS_LAYOUT.weekdayChart.col, buildWeekdayChartTable_(model)),
    monthlyCashflowChart: createSection_(ANALYTICS_LAYOUT.monthlyCashflowChart.row, ANALYTICS_LAYOUT.monthlyCashflowChart.col, buildMonthlyCashflowChartTable_(model)),
    topCategories: createSection_(ANALYTICS_LAYOUT.topCategories.row, ANALYTICS_LAYOUT.topCategories.col, buildTopListTable_('Category', model.categoryList)),
    topAccounts: createSection_(ANALYTICS_LAYOUT.topAccounts.row, ANALYTICS_LAYOUT.topAccounts.col, buildTopListTable_('Account', model.accountList)),
    topMerchants: createSection_(ANALYTICS_LAYOUT.topMerchants.row, ANALYTICS_LAYOUT.topMerchants.col, buildTopListTable_('Merchant', model.merchantList)),
    weekendSummary: createSection_(ANALYTICS_LAYOUT.weekendSummary.row, ANALYTICS_LAYOUT.weekendSummary.col, buildWeekendSummaryTable_(model)),
    topCategoriesChart: createSection_(ANALYTICS_LAYOUT.topCategoriesChart.row, ANALYTICS_LAYOUT.topCategoriesChart.col, buildChartTopListTable_('Category', model.categoryList, 22)),
    topAccountsChart: createSection_(ANALYTICS_LAYOUT.topAccountsChart.row, ANALYTICS_LAYOUT.topAccountsChart.col, buildChartTopListTable_('Account', model.accountList, 18)),
    topMerchantsChart: createSection_(ANALYTICS_LAYOUT.topMerchantsChart.row, ANALYTICS_LAYOUT.topMerchantsChart.col, buildChartTopListTable_('Merchant', model.merchantList, 20)),
    monthlyCategoryMatrix: createSection_(ANALYTICS_LAYOUT.monthlyCategoryMatrix.row, ANALYTICS_LAYOUT.monthlyCategoryMatrix.col, buildMonthlyMatrixTable_(model, model.topCategoryNames, 'categories')),
    monthlyAccountMatrix: createSection_(ANALYTICS_LAYOUT.monthlyAccountMatrix.row, ANALYTICS_LAYOUT.monthlyAccountMatrix.col, buildMonthlyMatrixTable_(model, model.topAccountNames, 'accounts')),
    weekendMonthlyCompare: createSection_(ANALYTICS_LAYOUT.weekendMonthlyCompare.row, ANALYTICS_LAYOUT.weekendMonthlyCompare.col, buildWeekendMonthlyCompareTable_(model)),
    weekendMonthlyCompareChart: createSection_(ANALYTICS_LAYOUT.weekendMonthlyCompareChart.row, ANALYTICS_LAYOUT.weekendMonthlyCompareChart.col, buildWeekendMonthlyCompareTable_(model)),
    categoryDriftChart: createSection_(ANALYTICS_LAYOUT.categoryDriftChart.row, ANALYTICS_LAYOUT.categoryDriftChart.col, buildCategoryDriftChartTable_(model)),
    anomalies: createSection_(ANALYTICS_LAYOUT.anomalies.row, ANALYTICS_LAYOUT.anomalies.col, buildAnomalyTable_(model)),
    recurring: createSection_(ANALYTICS_LAYOUT.recurring.row, ANALYTICS_LAYOUT.recurring.col, buildRecurringTable_(model)),
    categoryDrift: createSection_(ANALYTICS_LAYOUT.categoryDrift.row, ANALYTICS_LAYOUT.categoryDrift.col, buildCategoryDriftTable_(model))
  };
}

function writeAnalyticsSections_(sheet, sections) {
  sheet.clear();
  clearCharts_(sheet);
  Object.keys(sections).forEach(function(key) {
    const section = sections[key];
    writeTable_(sheet, section.row, section.col, section.values, '#161b22');
  });
  sheet.hideSheet();
}

function buildOverviewTable_(model) {
  return [
    ['Metric', 'Value'],
    ['Period Start', formatDateForPrompt_(model.minDate)],
    ['Period End', formatDateForPrompt_(model.maxDate)],
    ['Transactions', model.transactionCount],
    ['Expenses', model.expenseCount],
    ['Pending', model.pendingCount],
    ['External Spend', roundCurrency_(model.totalSpend)],
    ['External Income', roundCurrency_(model.totalIncome)],
    ['Net Cashflow', roundCurrency_(model.netCashflow)],
    ['Savings Rate', formatSavingsRateForDisplay_(model)],
    ['Daily Avg Burn', roundCurrency_(model.dailyAverageBurn)],
    ['Excluded By Rules Outflow', roundCurrency_(model.ruleExcludedCashOutflow)],
    ['Excluded By Rules Inflow', roundCurrency_(model.ruleExcludedCashInflow)],
    ['Excluded By Rules Count', model.ruleExcludedTransactionCount],
    ['Excluded Internal Outflow', roundCurrency_(model.excludedCashOutflow)],
    ['Excluded Internal Inflow', roundCurrency_(model.excludedCashInflow)],
    ['Excluded Internal Count', model.excludedTransactionCount],
    ['Review Only Count', model.reviewOnlyTransactionCount],
    ['Weekend Share', model.totalSpend > 0 ? model.weekendSpend / model.totalSpend : 0],
    ['Recurring Merchants', model.recurringCandidates.length]
  ];
}

function buildMonthlySummaryTable_(model) {
  const table = [['Month', 'Income', 'Spend', 'Net', 'Excluded Outflow', 'Excluded Inflow', 'Top Account', 'Top Category', 'Example']];
  model.monthKeys.forEach(function(monthKey) {
    const bucket = model.months[monthKey];
    table.push([
      formatMonthDisplayLabel_(monthKey),
      roundCurrency_(bucket.income),
      roundCurrency_(bucket.spend),
      roundCurrency_(bucket.net),
      roundCurrency_(bucket.excludedOutflow || 0),
      roundCurrency_(bucket.excludedInflow || 0),
      bucket.topAccount,
      bucket.topCategory,
      bucket.examples.length ? buildExampleList_(bucket.examples, 1) : 'N/A'
    ]);
  });
  if (table.length === 1) {
    table.push(['No data', 0, 0, 0, 0, 0, 'N/A', 'N/A', 'N/A']);
  }
  return table;
}

function buildWeekdaySummaryTable_(model) {
  const table = [['Day', 'Spend']];
  WEEKDAY_ORDER.forEach(function(day) {
    table.push([day, roundCurrency_(model.weekdays[day] || 0)]);
  });
  return table;
}

function buildWeeklySummaryTable_(model) {
  const table = [['Week Start', 'Spend', 'Income', 'Net']];
  model.weekKeys.forEach(function(weekKey) {
    const bucket = model.weeks[weekKey];
    table.push([
      weekKey,
      roundCurrency_(bucket.spend),
      roundCurrency_(bucket.income),
      roundCurrency_(bucket.net)
    ]);
  });
  if (table.length === 1) {
    table.push(['No data', 0, 0, 0]);
  }
  return table;
}

function buildWeekendSummaryTable_(model) {
  return [
    ['Metric', 'Value'],
    ['Weekend Spend', roundCurrency_(model.weekendSpend)],
    ['Weekday Spend', roundCurrency_(model.weekdaySpendTotal)],
    ['Weekend Share', model.totalSpend > 0 ? model.weekendSpend / model.totalSpend : 0],
    ['Top Weekend Category', model.weekendCategoryList.length ? model.weekendCategoryList[0].name : 'N/A'],
    ['Top Weekend Merchant', model.weekendMerchantList.length ? model.weekendMerchantList[0].name : 'N/A']
  ];
}

function buildMonthlyMatrixTable_(model, names, bucketKey) {
  const safeNames = names && names.length ? names : [bucketKey === 'accounts' ? 'Unknown Account' : 'Uncategorized'];
  const header = ['Month'].concat(safeNames.map(function(name) {
    return truncateLabel_(name, bucketKey === 'accounts' ? 14 : 18);
  }));
  const table = [header];
  model.monthKeys.forEach(function(monthKey) {
    const bucket = model.months[monthKey];
    const row = [formatMonthDisplayLabel_(monthKey)];
    safeNames.forEach(function(name) {
      row.push(roundCurrency_((bucket[bucketKey] && bucket[bucketKey][name]) || 0));
    });
    table.push(row);
  });
  if (table.length === 1) {
    table.push(['No data'].concat(safeNames.map(function() { return 0; })));
  }
  return table;
}

function buildWeekendMonthlyCompareTable_(model) {
  const table = [['Month', 'Weekday Spend', 'Weekend Spend']];
  model.monthKeys.forEach(function(monthKey) {
    const bucket = model.months[monthKey];
    table.push([
      formatMonthDisplayLabel_(monthKey),
      roundCurrency_(bucket.weekdaySpend || 0),
      roundCurrency_(bucket.weekendSpend || 0)
    ]);
  });
  if (table.length === 1) {
    table.push(['No data', 0, 0]);
  }
  return table;
}

function buildAnomalyTable_(model) {
  const table = [['Date', 'Merchant', 'Category', 'Spend', 'Account', 'Transaction ID']];
  model.anomalies.forEach(function(item) {
    table.push([item.date, item.merchant, item.category, roundCurrency_(item.spend), item.account, item.id]);
  });
  if (table.length === 1) {
    table.push(['No data', 'N/A', 'N/A', 0, 'N/A', 'N/A']);
  }
  return table;
}

function buildRecurringTable_(model) {
  const table = [['Merchant', 'Transactions', 'Total Spend', 'Avg Spend', 'Last Seen']];
  model.recurringCandidates.forEach(function(item) {
    table.push([
      item.name,
      item.count,
      roundCurrency_(item.total),
      roundCurrency_(item.count ? item.total / item.count : 0),
      item.lastSeen || 'N/A'
    ]);
  });
  if (table.length === 1) {
    table.push(['No recurring candidates', 0, 0, 0, 'N/A']);
  }
  return table;
}

function buildCategoryDriftTable_(model) {
  const table = [['Category', 'Latest Month', 'Prior Month', 'Delta']];
  model.categoryDrift.forEach(function(item) {
    table.push([
      item.name,
      roundCurrency_(item.latest),
      roundCurrency_(item.previous),
      roundCurrency_(item.delta)
    ]);
  });
  if (table.length === 1) {
    table.push(['No drift data', 0, 0, 0]);
  }
  return table;
}

function buildCategoryDriftChartTable_(model) {
  const table = [['Category', 'Delta']];
  model.categoryDrift.forEach(function(item) {
    table.push([truncateLabel_(item.name, 22), roundCurrency_(item.delta)]);
  });
  if (table.length === 1) {
    table.push(['No drift data', 0]);
  }
  return table;
}

function buildTopListTable_(label, list) {
  const table = [[label, 'Spend']];
  list.slice(0, 10).forEach(function(item) {
    table.push([item.name, roundCurrency_(item.total)]);
  });
  if (table.length === 1) {
    table.push(['No data', 0]);
  }
  return table;
}

function buildAnnotatedTopListTable_(label, list, truncateLimit) {
  const table = [[label, 'Spend', 'annotation']];
  list.slice(0, 8).forEach(function(item) {
    table.push([
      truncateLabel_(item.name, truncateLimit || 22),
      roundCurrency_(item.total),
      formatCurrency_(item.total)
    ]);
  });
  if (table.length === 1) {
    table.push(['No data', 0, formatCurrency_(0)]);
  }
  return table;
}

function buildChartTopListTable_(label, list, truncateLimit) {
  const table = [[label, 'Spend']];
  list.slice(0, 8).forEach(function(item) {
    table.push([
      truncateLabel_(item.name, truncateLimit || 22),
      roundCurrency_(item.total)
    ]);
  });
  if (table.length === 1) {
    table.push(['No data', 0]);
  }
  return table;
}

function buildAnnotatedWeekdayTable_(model) {
  const table = [['Day', 'Spend', 'annotation']];
  WEEKDAY_ORDER.forEach(function(day) {
    const spend = roundCurrency_(model.weekdays[day] || 0);
    table.push([day, spend, formatCurrency_(spend)]);
  });
  return table;
}

function buildWeekdayChartTable_(model) {
  const table = [['Day', 'Spend']];
  WEEKDAY_ORDER.forEach(function(day) {
    table.push([day, roundCurrency_(model.weekdays[day] || 0)]);
  });
  return table;
}

function buildAnnotatedWeekendMonthlyCompareTable_(model) {
  const table = [['Month', 'Weekday Spend', 'annotation', 'Weekend Spend', 'annotation']];
  model.monthKeys.forEach(function(monthKey) {
    const bucket = model.months[monthKey];
    const weekdaySpend = roundCurrency_(bucket.weekdaySpend || 0);
    const weekendSpend = roundCurrency_(bucket.weekendSpend || 0);
    table.push([
      formatMonthDisplayLabel_(monthKey),
      weekdaySpend,
      formatCurrency_(weekdaySpend),
      weekendSpend,
      formatCurrency_(weekendSpend)
    ]);
  });
  if (table.length === 1) {
    table.push(['No data', 0, formatCurrency_(0), 0, formatCurrency_(0)]);
  }
  return table;
}

function buildMonthlyCashflowChartTable_(model) {
  const hasIncomeSeries = monthlyCashflowHasIncome_(model);
  const table = hasIncomeSeries
    ? [['Month', 'Income', 'Spend', 'Net']]
    : [['Month', 'External Spend', 'Net Cashflow']];

  model.monthKeys.forEach(function(monthKey) {
    const bucket = model.months[monthKey];
    const monthLabel = formatMonthDisplayLabel_(monthKey);
    if (hasIncomeSeries) {
      table.push([
        monthLabel,
        roundCurrency_(bucket.income),
        roundCurrency_(bucket.spend),
        roundCurrency_(bucket.net)
      ]);
      return;
    }

    table.push([
      monthLabel,
      roundCurrency_(bucket.spend),
      roundCurrency_(bucket.net)
    ]);
  });

  if (table.length === 1) {
    table.push(hasIncomeSeries ? ['No data', 0, 0, 0] : ['No data', 0, 0]);
  }
  return table;
}

function monthlyCashflowHasIncome_(model) {
  return model.monthKeys.some(function(monthKey) {
    return roundCurrency_(model.months[monthKey].income || 0) !== 0;
  });
}

function buildSimpleAnnotationOptions_() {
  return {
    alwaysOutside: false,
    textStyle: {
      fontSize: 10,
      bold: true,
      auraColor: 'none',
      color: '#111827'
    }
  };
}

function buildAiContext_(model, records, query, intent, options) {
  options = options || {};
  const filters = extractRetrievalFilters_(model, query);
  const lines = [];
  lines.push('=== VERIFIED OVERVIEW ===');
  lines.push('Period -> ' + formatDateForPrompt_(model.minDate) + ' to ' + formatDateForPrompt_(model.maxDate) + ' (' + model.dayCount + ' days / ' + model.monthCount + ' months)');
  lines.push('Transactions -> ' + model.transactionCount + ' total, ' + model.expenseCount + ' expenses, ' + model.pendingCount + ' pending');
  lines.push('Totals -> External Spend ' + formatCurrency_(model.totalSpend) + ', External Income ' + formatCurrency_(model.totalIncome) + ', Net ' + formatCurrency_(model.netCashflow));
  lines.push('Excluded By Rules -> Outflow ' + formatCurrency_(model.ruleExcludedCashOutflow || 0) + ', Inflow ' + formatCurrency_(model.ruleExcludedCashInflow || 0) + ', Count ' + (model.ruleExcludedTransactionCount || 0));
  lines.push('Review-only Rule Flags -> ' + (model.reviewOnlyTransactionCount || 0) + ' transaction(s), still included in totals');
  lines.push('Excluded Internal Cash Movements -> Outflow ' + formatCurrency_(model.excludedCashOutflow) + ', Inflow ' + formatCurrency_(model.excludedCashInflow) + ', Count ' + model.excludedTransactionCount);
  if (model.activeRules && model.activeRules.length) {
    lines.push('Active Analytics Rules -> ' + summarizeAnalyticsRules_(model.activeRules));
  }
  const savingsStatus = buildSavingsRateStatus_(model);
  lines.push('Savings Rate -> ' + savingsStatus.status);
  lines.push('Income Coverage -> ' + savingsStatus.note);
  lines.push('Daily Burn -> ' + formatCurrency_(model.dailyAverageBurn));
  lines.push('Top Accounts -> [' + formatBucketSummary_(model.accountList, MAX_TOP_ITEMS) + ']');
  lines.push('Top Categories -> [' + formatBucketSummary_(model.categoryList, MAX_TOP_ITEMS) + ']');
  lines.push('Top Merchants -> [' + formatBucketSummary_(model.merchantList, MAX_TOP_ITEMS) + ']');

  lines.push('');
  lines.push('=== MONTHLY TOTALS ===');
  model.monthKeys.forEach(function(monthKey) {
    const bucket = model.months[monthKey];
    lines.push(monthKey + ' -> Spend ' + formatCurrency_(bucket.spend) + ' | Income ' + formatCurrency_(bucket.income) + ' | Net ' + formatCurrency_(bucket.net));
  });

  lines.push('');
  lines.push('=== WEEKDAY SPEND ===');
  WEEKDAY_ORDER.forEach(function(day) {
    lines.push(day + ' -> ' + formatCurrency_(model.weekdays[day] || 0));
  });
  lines.push('Weekend Total -> ' + formatCurrency_(model.weekendSpend));
  lines.push('Weekday Total -> ' + formatCurrency_(model.weekdaySpendTotal));
  lines.push('Weekend Share -> ' + formatPercent_(model.totalSpend > 0 ? model.weekendSpend / model.totalSpend : 0));

  if (intent.needsWeekly) {
    lines.push('');
    lines.push('=== WEEKLY TOTALS ===');
    model.weekKeys.forEach(function(weekKey) {
      const bucket = model.weeks[weekKey];
      lines.push(weekKey + ' -> Spend ' + formatCurrency_(bucket.spend) + ' | Income ' + formatCurrency_(bucket.income) + ' | Net ' + formatCurrency_(bucket.net));
    });
  }

  if (intent.needsMonthly) {
    lines.push('');
    lines.push('=== MONTHLY ACCOUNT BREAKDOWN ===');
    model.monthKeys.forEach(function(monthKey) {
      const bucket = model.months[monthKey];
      lines.push(monthKey + ' -> [' + formatBucketSummary_(bucket.accountList, Math.max(bucket.accountList.length, MAX_TOP_ITEMS)) + ']');
    });

    lines.push('');
    lines.push('=== MONTHLY CATEGORY BREAKDOWN ===');
    model.monthKeys.forEach(function(monthKey) {
      const bucket = model.months[monthKey];
      lines.push(monthKey + ' -> [' + formatBucketSummary_(bucket.categoryList, Math.max(bucket.categoryList.length, MAX_TOP_ITEMS)) + ']');
    });

    lines.push('');
    lines.push('=== MONTHLY EXAMPLES ===');
    model.monthKeys.forEach(function(monthKey) {
      const bucket = model.months[monthKey];
      lines.push(monthKey + ' -> [' + buildExampleList_(bucket.examples, MAX_CATEGORY_EXAMPLES) + ']');
    });

    lines.push('');
    lines.push('=== SCOPED MONTH DETAIL ===');
    buildScopedMonthDetailLines_(records, model, filters.months.length ? filters.months : model.monthKeys).forEach(function(line) {
      lines.push(line);
    });
  }

  if (intent.needsWeekend) {
    lines.push('');
    lines.push('=== WEEKEND LEAKS ===');
    lines.push('Top Weekend Categories -> [' + formatBucketSummary_(model.weekendCategoryList, 5) + ']');
    lines.push('Top Weekend Merchants -> [' + formatBucketSummary_(model.weekendMerchantList, 5) + ']');
    lines.push('Examples -> [' + buildExampleList_(model.weekendExamples, MAX_CATEGORY_EXAMPLES) + ']');
  }

  if (intent.needsAccount && !intent.needsMonthly) {
    lines.push('');
    lines.push('=== ACCOUNT BREAKDOWN ===');
    lines.push('Accounts -> [' + formatBucketSummary_(model.accountList, MAX_TOP_ITEMS) + ']');
  }

  if (intent.needsCategoryExamples) {
    const selectedCategories = selectDetailedCategories_(query, model.detailedCategoryList);
    lines.push('');
    lines.push('=== CATEGORY DETAIL BLOCK ===');
    selectedCategories.forEach(function(categoryInfo) {
      lines.push(categoryInfo.name + ' -> ' + formatCurrency_(categoryInfo.total) + ' -> [' + buildExampleList_(categoryInfo.examples, MAX_CATEGORY_EXAMPLES) + ']');
    });
  }

  if (intent.needsAnomalies) {
    lines.push('');
    lines.push('=== LARGEST TRANSACTIONS ===');
    model.anomalies.slice(0, 6).forEach(function(item) {
      lines.push(item.date + ' -> ' + item.merchant + ' -> ' + formatCurrency_(item.spend) + ' -> ' + item.account + ' -> (' + item.id + ')');
    });
  }

  if (!intent.needsStructuredReport || intent.needsAdvice) {
    const diagnostics = buildStrategistDiagnostics_(model);
    const latestMonthSnapshot = options.includeLatestMonthSnapshot === false ? [] : buildLatestMonthSnapshot_(model);
    const recurringContext = options.includeRecurringContext === false ? [] : buildRecurringContext_(model);
    const driftContext = options.includeDriftContext === false ? [] : buildDriftContext_(model);
    const recentContext = options.includeRecentContext === false ? [] : buildRecentTransactionContext_(records);
    const retrievalSummary = buildRetrievalSummary_(records, filters);
    const evidence = selectRelevantTransactions_(records, query, intent, filters, options.evidenceLimit);

    lines.push('');
    lines.push('=== STRATEGIST DIAGNOSTICS ===');
    diagnostics.forEach(function(line) {
      lines.push(line);
    });

    if (latestMonthSnapshot.length) {
      lines.push('');
      lines.push('=== LATEST MONTH SNAPSHOT ===');
      latestMonthSnapshot.forEach(function(line) {
        lines.push(line);
      });
    }

    if (recurringContext.length) {
      lines.push('');
      lines.push('=== RECURRING MERCHANT CONTEXT ===');
      recurringContext.forEach(function(line) {
        lines.push(line);
      });
    }

    if (driftContext.length) {
      lines.push('');
      lines.push('=== CATEGORY DRIFT CONTEXT ===');
      driftContext.forEach(function(line) {
        lines.push(line);
      });
    }

    if (retrievalSummary.length) {
      lines.push('');
      lines.push('=== QUERY RETRIEVAL SUMMARY ===');
      retrievalSummary.forEach(function(line) {
        lines.push(line);
      });
    }

    lines.push('');
    lines.push('=== RELEVANT TRANSACTION EVIDENCE ===');
    evidence.forEach(function(line) {
      lines.push(line);
    });

    if (recentContext.length) {
      lines.push('');
      lines.push('=== RECENT TRANSACTION SNAPSHOT ===');
      recentContext.forEach(function(line) {
        lines.push(line);
      });
    }
  }

  return lines.join('\n');
}

function buildConversationQuery_(conversationTurns, query) {
  const priorUserTurns = (conversationTurns || []).slice(-MAX_CONTEXT_TURNS).map(function(turn) {
    return String(turn.user || '').trim();
  }).filter(Boolean);
  return priorUserTurns.concat([String(query || '').trim()]).join('\n');
}

function buildConversationContents_(conversationTurns, query, backgroundContext) {
  const contents = [];
  (conversationTurns || []).slice(-MAX_CONTEXT_TURNS).forEach(function(turn) {
    if (turn.user) {
      contents.push({
        role: 'user',
        parts: [{ text: turn.user }]
      });
    }
    if (turn.assistant) {
      contents.push({
        role: 'model',
        parts: [{ text: turn.assistant }]
      });
    }
  });

  contents.push({
    role: 'user',
    parts: [{
      text: [
        'VERIFIED FINANCE CONTEXT',
        backgroundContext,
        '',
        'CURRENT USER REQUEST',
        String(query || '')
      ].join('\n')
    }]
  });
  return contents;
}

function buildGeminiBackgroundContext_(model, records, conversationQuery, intent, filters) {
  const lines = [];
  lines.push(buildAiContext_(model, records, conversationQuery, intent, {
    evidenceLimit: Math.max(MAX_EVIDENCE_TRANSACTIONS * 2, 24)
  }));
  lines.push('');
  lines.push(buildResponseContract_(intent));

  const rawLedger = buildRawLedgerContext_(records, filters);
  if (rawLedger.length) {
    lines.push('');
    lines.push('=== RAW TRANSACTION LEDGER ===');
    rawLedger.forEach(function(line) {
      lines.push(line);
    });
  }

  return lines.join('\n');
}

function buildRawLedgerContext_(records, filters) {
  const activeFilters = filters || { months: [], categories: [], merchants: [] };
  const scoped = hasActiveFilters_(activeFilters) ? filterTransactionsByRetrieval_(records, activeFilters) : [];
  let sourceRecords = scoped.length ? scoped : records.slice();
  sourceRecords = sourceRecords.filter(function(record) {
    return record.date;
  }).sort(function(a, b) {
    return b.date.getTime() - a.date.getTime();
  });

  const maxRows = sourceRecords.length <= FULL_LEDGER_TRANSACTION_THRESHOLD
    ? sourceRecords.length
    : MAX_LEDGER_CONTEXT_TRANSACTIONS;

  return sourceRecords.slice(0, maxRows).map(function(record) {
    const amount = Number(record.amount || 0);
    const signedAmount = amount >= 0
      ? '+' + formatCurrency_(amount)
      : '-' + formatCurrency_(Math.abs(amount));
    return [
      formatDateForPrompt_(record.date),
      truncateLabel_(record.name, 42),
      signedAmount,
      formatDetailedCategoryLabel_(record.category),
      truncateLabel_(formatAccountLabel_(record.account), 22),
      record.pending ? 'pending' : 'posted',
      '(' + record.id + ')'
    ].join(' | ');
  });
}

function buildResponseContract_(intent) {
  const lines = ['=== RESPONSE CONTRACT ==='];
  lines.push('Use only the verified blocks that were provided.');
  lines.push('If a requested field is not included in the verified blocks, say that it is not included instead of guessing.');

  if (intent.needsMonthly) {
    lines.push('For monthly questions, include every non-zero category in scope.');
    lines.push('For monthly questions, respond with this exact shape:');
    lines.push('1. A markdown table of categories and spend for each resolved month.');
    lines.push('2. A full account breakdown for each resolved month.');
    lines.push('3. A section named "Grouped Transactions By Category" with category -> sample transactions.');
    lines.push('4. If advice was requested, sections named Quick Wins, Subscriptions, Behavior Patterns, and Watch List.');
  }
  if (intent.needsWeekend) {
    lines.push('For weekend questions, respond as: Weekend total -> Weekend share -> Categories[...] -> Merchants[...] -> Examples[...]');
  }
  if (intent.needsCategoryExamples) {
    lines.push('For category detail questions, respond as: Category -> Total -> [Merchant (Transaction ID)]');
  }
  if (!intent.needsMonthly && !intent.needsWeekend && !intent.needsCategoryExamples) {
    lines.push('For overview questions, answer in 3-5 compact bullets or short paragraphs.');
  }
  if (intent.needsCashflow) {
    lines.push('For savings-rate questions, include the verified external income amount, excluded payment/transfer amount, and whether the savings rate is calculable.');
    lines.push('Do not describe 0% as the true savings rate when verified external income is $0.00.');
  }

  return lines.join('\n');
}

function buildDirectAnswer_(model, query, intent, filters) {
  if (!intent.needsStructuredReport) {
    return '';
  }

  const lines = [];
  if (hasActiveFilters_(filters)) {
    lines.push('Scope -> ' + buildFilterLabel_(filters));
    lines.push('');
  }

  if (intent.needsMonthly) {
    lines.push('Here is your verified monthly spend breakdown:');
    model.monthKeys.forEach(function(monthKey) {
      const bucket = model.months[monthKey];
      const monthRecords = filterRecordsByToolArgs_(getTransactionRecords_(), model, {
        month: monthKey,
        expensesOnly: true
      });
      lines.push('');
      lines.push(monthKey);
      lines.push('Spend -> ' + formatCurrency_(bucket.spend));
      lines.push('Income -> ' + formatCurrency_(bucket.income));
      lines.push('Net -> ' + formatCurrency_(bucket.net));
      lines.push('Excluded By Rules Outflow -> ' + formatCurrency_(bucket.ruleExcludedOutflow || 0));
      lines.push('Excluded By Rules Inflow -> ' + formatCurrency_(bucket.ruleExcludedInflow || 0));
      lines.push('Excluded Internal Outflow -> ' + formatCurrency_(bucket.excludedOutflow || 0));
      lines.push('Excluded Internal Inflow -> ' + formatCurrency_(bucket.excludedInflow || 0));
      lines.push('Accounts -> [' + formatBucketSummary_(bucket.accountList, Math.max(bucket.accountList.length, MAX_TOP_ITEMS)) + ']');
      lines.push('Categories -> [' + formatBucketSummary_(bucket.categoryList, Math.max(bucket.categoryList.length, MAX_TOP_ITEMS)) + ']');
      lines.push('Grouped Transactions By Category ->');
      buildGroupedCategoryExampleLines_(monthRecords).forEach(function(line) {
        lines.push(line);
      });
      lines.push('Examples -> [' + buildExampleList_(bucket.examples, MAX_CATEGORY_EXAMPLES) + ']');
    });
  }

  if (intent.needsWeekend) {
    const weekendLeader = model.weekendCategoryList.length ? model.weekendCategoryList[0].name + ' ' + formatCurrency_(model.weekendCategoryList[0].total) : 'None';
    const merchantLeader = model.weekendMerchantList.length ? model.weekendMerchantList[0].name + ' ' + formatCurrency_(model.weekendMerchantList[0].total) : 'None';
    if (lines.length) {
      lines.push('');
    }
    lines.push('Weekend leak summary:');
    lines.push('Weekend Spend -> ' + formatCurrency_(model.weekendSpend));
    lines.push('Weekday Spend -> ' + formatCurrency_(model.weekdaySpendTotal));
    lines.push('Weekend Share -> ' + formatPercent_(model.totalSpend > 0 ? model.weekendSpend / model.totalSpend : 0));
    lines.push('Top Weekend Category -> ' + weekendLeader);
    lines.push('Top Weekend Merchant -> ' + merchantLeader);
    lines.push('Weekend Examples -> [' + buildExampleList_(model.weekendExamples, MAX_CATEGORY_EXAMPLES) + ']');
    lines.push('Weekend by Month -> [' + model.monthKeys.map(function(monthKey) {
      const bucket = model.months[monthKey];
      return monthKey + ' ' + formatCurrency_(bucket.weekendSpend || 0);
    }).join(', ') + ']');
  }

  if (intent.needsCategoryExamples && !intent.needsMonthly) {
    const selectedCategories = selectDetailedCategories_(query, model.detailedCategoryList);
    if (lines.length) {
      lines.push('');
    }
    lines.push('Category detail:');
    selectedCategories.forEach(function(categoryInfo) {
      lines.push(categoryInfo.name + ' -> ' + formatCurrency_(categoryInfo.total) + ' -> [' + buildExampleList_(categoryInfo.examples, MAX_CATEGORY_EXAMPLES) + ']');
    });
  }

  if (intent.needsAnomalies) {
    if (lines.length) {
      lines.push('');
    }
    lines.push('Largest transactions:');
    model.anomalies.slice(0, 8).forEach(function(item) {
      lines.push(item.date + ' -> ' + item.merchant + ' -> ' + formatCurrency_(item.spend) + ' -> ' + item.account + ' -> (' + item.id + ')');
    });
  }

  if (intent.needsAdvice) {
    appendAdviceSections_(lines, model);
  }

  return lines.join('\n');
}

function buildHeuristicObservations_(model, intent) {
  const observations = [];
  const weekendShare = model.totalSpend > 0 ? model.weekendSpend / model.totalSpend : 0;
  if (intent.needsWeekend && weekendShare >= 0.3) {
    observations.push('Weekend spend is a meaningful share of total spend at ' + formatPercent_(weekendShare) + '.');
  }
  if (model.categoryList.length) {
    observations.push('The largest spending category is ' + model.categoryList[0].name + ' at ' + formatCurrency_(model.categoryList[0].total) + '.');
  }
  if (model.merchantList.length) {
    observations.push('The largest merchant concentration is ' + model.merchantList[0].name + ' at ' + formatCurrency_(model.merchantList[0].total) + '.');
  }
  return observations.slice(0, 3);
}

function buildStrategistDiagnostics_(model) {
  const lines = [];
  const averageMonthlySpend = model.monthKeys.length
    ? model.totalSpend / Math.max(1, model.monthKeys.length)
    : model.totalSpend;
  const topCategoryShare = model.categoryList.length && model.totalSpend
    ? model.categoryList[0].total / model.totalSpend
    : 0;
  const topMerchantShare = model.merchantList.length && model.totalSpend
    ? model.merchantList[0].total / model.totalSpend
    : 0;
  const recurringTotal = model.recurringCandidates.reduce(function(sum, item) {
    return sum + item.total;
  }, 0);
  const highestMonth = getExtremeMonth_(model, 'max');
  const lowestMonth = getExtremeMonth_(model, 'min');
  const latestMonth = model.monthKeys.length ? model.monthKeys[model.monthKeys.length - 1] : null;
  const previousMonth = model.monthKeys.length > 1 ? model.monthKeys[model.monthKeys.length - 2] : null;
  const latestBucket = latestMonth ? model.months[latestMonth] : null;
  const previousBucket = previousMonth ? model.months[previousMonth] : null;
  const monthDelta = latestBucket && previousBucket ? latestBucket.spend - previousBucket.spend : 0;
  const biggestDrift = model.categoryDrift.length ? model.categoryDrift[0] : null;

    lines.push('Average Monthly Spend -> ' + formatCurrency_(averageMonthlySpend));
  if (highestMonth) {
    lines.push('Highest Spend Month -> ' + highestMonth.key + ' ' + formatCurrency_(highestMonth.spend));
  }
  if (lowestMonth) {
    lines.push('Lowest Spend Month -> ' + lowestMonth.key + ' ' + formatCurrency_(lowestMonth.spend));
  }
  if (latestBucket) {
    lines.push('Latest Month Spend -> ' + latestMonth + ' ' + formatCurrency_(latestBucket.spend));
  }
  if (latestBucket && previousBucket) {
    lines.push('Latest vs Prior Month Delta -> ' + formatCurrency_(monthDelta));
  }
  if (model.categoryList.length) {
    lines.push('Top Category Concentration -> ' + model.categoryList[0].name + ' ' + formatPercent_(topCategoryShare));
  }
  if (model.merchantList.length) {
    lines.push('Top Merchant Concentration -> ' + model.merchantList[0].name + ' ' + formatPercent_(topMerchantShare));
  }
  if (recurringTotal > 0) {
    lines.push('Recurring Merchant Spend -> ' + formatCurrency_(recurringTotal) + ' across ' + model.recurringCandidates.length + ' merchants');
  }
  if (biggestDrift) {
    lines.push('Largest Category Drift -> ' + biggestDrift.name + ' ' + formatCurrency_(biggestDrift.delta));
  }

  return lines;
}

function buildLatestMonthSnapshot_(model) {
  if (!model.monthKeys.length) {
    return [];
  }

  const latestKey = model.monthKeys[model.monthKeys.length - 1];
  const bucket = model.months[latestKey];
  return [
    'Month -> ' + latestKey,
    'Spend -> ' + formatCurrency_(bucket.spend),
    'Income -> ' + formatCurrency_(bucket.income),
    'Excluded By Rules Outflow -> ' + formatCurrency_(bucket.ruleExcludedOutflow || 0),
    'Excluded By Rules Inflow -> ' + formatCurrency_(bucket.ruleExcludedInflow || 0),
    'Excluded Internal Outflow -> ' + formatCurrency_(bucket.excludedOutflow || 0),
    'Excluded Internal Inflow -> ' + formatCurrency_(bucket.excludedInflow || 0),
    'Accounts -> [' + formatBucketSummary_(bucket.accountList, Math.max(bucket.accountList.length, 6)) + ']',
    'Categories -> [' + formatBucketSummary_(bucket.categoryList, Math.max(bucket.categoryList.length, 8)) + ']',
    'Merchants -> [' + formatBucketSummary_(bucket.merchantList, 6) + ']',
    'Examples -> [' + buildExampleList_(bucket.examples, MAX_CATEGORY_EXAMPLES) + ']'
  ];
}

function buildRecurringContext_(model) {
  if (!model.recurringCandidates.length) {
    return [];
  }

  return model.recurringCandidates.slice(0, 8).map(function(item) {
    const average = item.count ? item.total / item.count : 0;
    return item.name + ' -> Total ' + formatCurrency_(item.total) +
      ' -> Count ' + item.count +
      ' -> Avg ' + formatCurrency_(average) +
      ' -> Last Seen ' + (item.lastSeen || 'Unknown');
  });
}

function buildDriftContext_(model) {
  if (!model.categoryDrift.length) {
    return [];
  }

  return model.categoryDrift.slice(0, 8).map(function(item) {
    return item.name + ' -> Latest ' + formatCurrency_(item.latest) +
      ' -> Prior ' + formatCurrency_(item.previous) +
      ' -> Delta ' + formatCurrency_(item.delta);
  });
}

function buildRecentTransactionContext_(records) {
  return records
    .filter(function(record) {
      return record.date && Number(record.amount || 0) < 0;
    })
    .sort(function(a, b) {
      return b.date.getTime() - a.date.getTime();
    })
    .slice(0, 8)
    .map(function(record) {
      return formatDateForPrompt_(record.date) + ' -> ' +
        truncateLabel_(record.name, 28) + ' -> ' +
        formatDetailedCategoryLabel_(record.category) + ' -> ' +
        formatCurrency_(Math.abs(record.amount)) + ' -> (' + record.id + ')';
    });
}

function buildRetrievalSummary_(records, filters) {
  const filtered = filterTransactionsByRetrieval_(records, filters);
  const expenses = filtered.filter(function(record) {
    return Number(record.amount || 0) < 0;
  });
  const spend = expenses.reduce(function(sum, record) {
    return sum + Math.abs(Number(record.amount || 0));
  }, 0);
  const categoryTotals = {};
  const merchantTotals = {};
  const accountTotals = {};

  expenses.forEach(function(record) {
    const category = formatDetailedCategoryLabel_(record.category);
    const merchant = truncateLabel_(record.name, 36);
    const account = formatAccountLabel_(record.account);
    const amount = Math.abs(Number(record.amount || 0));
    categoryTotals[category] = (categoryTotals[category] || 0) + amount;
    merchantTotals[merchant] = (merchantTotals[merchant] || 0) + amount;
    accountTotals[account] = (accountTotals[account] || 0) + amount;
  });

  const lines = [];
  if (filters.months.length) {
    lines.push('Months -> [' + filters.months.join(', ') + ']');
  }
  if (filters.categories.length) {
    lines.push('Categories -> [' + filters.categories.join(', ') + ']');
  }
  if (filters.merchants.length) {
    lines.push('Merchants -> [' + filters.merchants.join(', ') + ']');
  }
  if (filters.accounts && filters.accounts.length) {
    lines.push('Accounts -> [' + filters.accounts.join(', ') + ']');
  }
  lines.push('Matched Transactions -> ' + filtered.length);
  lines.push('Matched Spend -> ' + formatCurrency_(spend));
  lines.push('Matched Accounts -> [' + formatBucketSummary_(bucketMapToSortedList_(accountTotals), 5) + ']');
  lines.push('Matched Categories -> [' + formatBucketSummary_(bucketMapToSortedList_(categoryTotals), 5) + ']');
  lines.push('Matched Merchants -> [' + formatBucketSummary_(bucketMapToSortedList_(merchantTotals), 5) + ']');
  return lines;
}

function getExtremeMonth_(model, mode) {
  if (!model.monthKeys.length) {
    return null;
  }

  return model.monthKeys.map(function(monthKey) {
    return {
      key: monthKey,
      spend: model.months[monthKey].spend
    };
  }).sort(function(a, b) {
    return mode === 'min' ? a.spend - b.spend : b.spend - a.spend;
  })[0];
}

function extractRetrievalFilters_(model, query) {
  const normalized = normalizeQueryText_(query);
  return {
    months: extractMentionedMonths_(model, normalized),
    categories: extractMentionedCategories_(model, normalized),
    merchants: extractMentionedMerchants_(model, normalized),
    accounts: extractMentionedAccounts_(model, normalized)
  };
}

function extractMentionedMonths_(model, normalizedQuery) {
  const months = [];
  const relativeMonths = resolveRelativeMonths_(model, normalizedQuery);
  relativeMonths.forEach(function(monthKey) {
    months.push(monthKey);
  });

  const isoMatches = normalizedQuery.match(/\b20\d{2}\s(0[1-9]|1[0-2])\b/g) || [];
  isoMatches.forEach(function(match) {
    months.push(match.replace(' ', '-'));
  });

  const monthAliases = [
    ['january', 'jan'],
    ['february', 'feb'],
    ['march', 'mar'],
    ['april', 'apr'],
    ['may'],
    ['june', 'jun'],
    ['july', 'jul'],
    ['august', 'aug'],
    ['september', 'sep', 'sept'],
    ['october', 'oct'],
    ['november', 'nov'],
    ['december', 'dec']
  ];
  monthAliases.forEach(function(aliases, index) {
    const found = aliases.some(function(alias) {
      return new RegExp('(^| )' + alias + '( |$)').test(normalizedQuery);
    });
    if (!found) {
      return;
    }
    const monthNumber = ('0' + (index + 1)).slice(-2);
    const yearMatches = normalizedQuery.match(/\b20\d{2}\b/g) || [];
    if (yearMatches.length) {
      yearMatches.forEach(function(year) {
        months.push(year + '-' + monthNumber);
      });
    } else {
      model.monthKeys.forEach(function(monthKey) {
        if (monthKey.slice(5, 7) === monthNumber) {
          months.push(monthKey);
        }
      });
    }
  });

  return months.filter(uniqueValue_);
}

function resolveRelativeMonths_(model, normalizedQuery) {
  if (!model || !model.monthKeys || !model.monthKeys.length) {
    return [];
  }

  const resolved = [];
  const latestMonth = model.monthKeys[model.monthKeys.length - 1];
  const previousMonth = model.monthKeys.length > 1 ? model.monthKeys[model.monthKeys.length - 2] : null;
  const currentCalendarMonth = Utilities.formatDate(new Date(), getSpreadsheetTimeZone_(), 'yyyy-MM');
  const currentDatasetMonth = model.monthKeys.indexOf(currentCalendarMonth) !== -1 ? currentCalendarMonth : latestMonth;

  if (/(this month|current month|this months|current months)/.test(normalizedQuery)) {
    resolved.push(currentDatasetMonth);
  }
  if (/(last month|previous month|prior month)/.test(normalizedQuery)) {
    resolved.push(previousMonth || currentDatasetMonth);
  }
  if (/(latest month|most recent month)/.test(normalizedQuery)) {
    resolved.push(latestMonth);
  }

  return resolved.filter(Boolean);
}

function hasActiveFilters_(filters) {
  return Boolean(filters && (
    (filters.months && filters.months.length) ||
    (filters.categories && filters.categories.length) ||
    (filters.merchants && filters.merchants.length) ||
    (filters.accounts && filters.accounts.length)
  ));
}

function buildFilterLabel_(filters) {
  const parts = [];
  if (filters.months && filters.months.length) {
    parts.push('Months [' + filters.months.join(', ') + ']');
  }
  if (filters.categories && filters.categories.length) {
    parts.push('Categories [' + filters.categories.join(', ') + ']');
  }
  if (filters.merchants && filters.merchants.length) {
    parts.push('Merchants [' + filters.merchants.join(', ') + ']');
  }
  if (filters.accounts && filters.accounts.length) {
    parts.push('Accounts [' + filters.accounts.join(', ') + ']');
  }
  return parts.length ? parts.join(' | ') : 'Entire dataset';
}

function buildOverviewToolResult_(model) {
  const savingsStatus = buildSavingsRateStatus_(model);
  return {
    period_start: formatDateForPrompt_(model.minDate),
    period_end: formatDateForPrompt_(model.maxDate),
    total_spend: roundCurrency_(model.totalSpend),
    total_income: roundCurrency_(model.totalIncome),
    excluded_by_rules_outflow: roundCurrency_(model.ruleExcludedCashOutflow || 0),
    excluded_by_rules_inflow: roundCurrency_(model.ruleExcludedCashInflow || 0),
    excluded_by_rules_count: model.ruleExcludedTransactionCount || 0,
    excluded_internal_outflow: roundCurrency_(model.excludedCashOutflow),
    excluded_internal_inflow: roundCurrency_(model.excludedCashInflow),
    excluded_internal_count: model.excludedTransactionCount,
    review_only_count: model.reviewOnlyTransactionCount || 0,
    net_cashflow: roundCurrency_(model.netCashflow),
    savings_rate: roundCurrency_(model.savingsRate * 100),
    savings_rate_status: savingsStatus.status,
    income_coverage_note: savingsStatus.note,
    daily_average_burn: roundCurrency_(model.dailyAverageBurn),
    top_accounts: serializeNamedTotals_(model.accountList, 6),
    top_categories: serializeNamedTotals_(model.categoryList, 6),
    top_merchants: serializeMerchantItems_(model.merchantList, 6),
    latest_month: buildLatestMonthSnapshot_(model),
    recurring_merchants: serializeMerchantItems_(model.recurringCandidates, 6),
    category_drift: model.categoryDrift.slice(0, 6)
  };
}

function buildSavingsRateStatus_(model) {
  if (!model || Number(model.totalIncome || 0) <= 0) {
    return {
      status: 'Not calculable from current linked accounts',
      note: 'Verified external income is $0.00 after excluding credit-card payments/transfers. Link a checking/payroll account or import income rows before treating savings rate as real.'
    };
  }
  return {
    status: formatPercent_(model.savingsRate),
    note: 'Verified external income is present, so savings rate is calculated as net cashflow divided by external income.'
  };
}

function formatSavingsRateForDisplay_(model) {
  if (!model || Number(model.totalIncome || 0) <= 0) {
    return 'N/A - no verified income';
  }
  return formatPercent_(model.savingsRate);
}

function buildMonthBreakdownToolResult_(records, model, args) {
  const filtered = filterRecordsByToolArgs_(records, model, {
    month: args.month,
    expensesOnly: false
  });
  const sourceRecords = filtered.length ? filtered : (args.month ? [] : records);
  const scopedModel = buildAnalyticsModel_(sourceRecords);
  return {
    scope: buildToolScope_(sourceRecords, model, args),
    months: scopedModel.monthKeys.map(function(monthKey) {
      const bucket = scopedModel.months[monthKey];
      const monthRecords = filterRecordsByToolArgs_(sourceRecords, scopedModel, {
        month: monthKey,
        expensesOnly: true
      });
      return {
        month: monthKey,
        month_label: formatMonthDisplayLabel_(monthKey),
        spend: roundCurrency_(bucket.spend),
        income: roundCurrency_(bucket.income),
        net: roundCurrency_(bucket.net),
        excluded_internal_outflow: roundCurrency_(bucket.excludedOutflow || 0),
        excluded_internal_inflow: roundCurrency_(bucket.excludedInflow || 0),
        accounts: args.include_accounts === false ? [] : serializeNamedTotals_(bucket.accountList, MAX_FULL_BREAKDOWN_ITEMS),
        categories: args.include_categories === false ? [] : serializeNamedTotals_(bucket.categoryList, MAX_FULL_BREAKDOWN_ITEMS),
        transactions_by_category: args.include_examples === false ? [] : serializeTransactionsByCategory_(monthRecords, MAX_GROUPED_CATEGORY_EXAMPLES),
        example_transactions: args.include_examples === false ? [] : serializeTransactions_(monthRecords, 8)
      };
    })
  };
}

function buildCategoryBreakdownToolResult_(records, model, args) {
  const filtered = filterRecordsByToolArgs_(records, model, {
    month: args.month,
    category: args.category,
    expensesOnly: true
  });
  const scopedModel = buildAnalyticsModel_(filtered);
  return {
    scope: buildToolScope_(filtered, model, args),
    total_spend: roundCurrency_(scopedModel.totalSpend),
    accounts: serializeNamedTotals_(scopedModel.accountList, 8),
    categories: serializeNamedTotals_(scopedModel.detailedCategoryList, 8),
    example_transactions: args.include_examples === false ? [] : serializeTransactions_(filtered, 6)
  };
}

function buildAccountBreakdownToolResult_(records, model, args) {
  const filtered = filterRecordsByToolArgs_(records, model, {
    month: args.month,
    account: args.account,
    expensesOnly: true
  });
  const scopedModel = buildAnalyticsModel_(filtered);
  return {
    scope: buildToolScope_(filtered, model, args),
    total_spend: roundCurrency_(scopedModel.totalSpend),
    categories: serializeNamedTotals_(scopedModel.categoryList, 8),
    merchants: serializeMerchantItems_(scopedModel.merchantList, 8),
    example_transactions: args.include_examples === false ? [] : serializeTransactions_(filtered, 6)
  };
}

function buildWeekendAnalysisToolResult_(records, model, args) {
  const filtered = filterRecordsByToolArgs_(records, model, {
    month: args.month,
    expensesOnly: true
  });
  const sourceRecords = filtered.length ? filtered : (args.month ? [] : records);
  const scopedModel = buildAnalyticsModel_(sourceRecords);
  return {
    scope: buildToolScope_(sourceRecords, model, args),
    weekend_spend: roundCurrency_(scopedModel.weekendSpend),
    weekday_spend: roundCurrency_(scopedModel.weekdaySpendTotal),
    weekend_share: roundCurrency_((scopedModel.totalSpend > 0 ? scopedModel.weekendSpend / scopedModel.totalSpend : 0) * 100),
    top_weekend_categories: serializeNamedTotals_(scopedModel.weekendCategoryList, 6),
    top_weekend_merchants: serializeNamedTotals_(scopedModel.weekendMerchantList, 6),
    example_transactions: args.include_examples === false ? [] : serializeTransactions_(sourceRecords.filter(function(record) {
      return record.date && isWeekendDate_(record.date) && Number(record.amount || 0) < 0;
    }), 6)
  };
}

function buildSearchTransactionsToolResult_(records, model, args) {
  const filtered = filterRecordsByToolArgs_(records, model, {
    month: args.month,
    category: args.category,
    merchant: args.merchant,
    account: args.account,
    expensesOnly: args.expenses_only === true
  });
  const sort = String(args.sort || 'largest').toLowerCase();
  const sorted = filtered.slice().sort(function(a, b) {
    if (sort === 'recent') {
      return (b.date ? b.date.getTime() : 0) - (a.date ? a.date.getTime() : 0);
    }
    return Math.abs(Number(b.amount || 0)) - Math.abs(Number(a.amount || 0));
  });
  const limited = sorted.slice(0, normalizeLimit_(args.limit, 8));
  return {
    scope: buildToolScope_(filtered, model, args),
    matched_spend: roundCurrency_(filtered.reduce(function(sum, record) {
      return sum + (Number(record.amount || 0) < 0 ? Math.abs(Number(record.amount || 0)) : 0);
    }, 0)),
    matched_income: roundCurrency_(filtered.reduce(function(sum, record) {
      return sum + (Number(record.amount || 0) > 0 ? Number(record.amount || 0) : 0);
    }, 0)),
    transactions: serializeTransactionsInOrder_(limited, limited.length || normalizeLimit_(args.limit, 8)),
    table_headers: ['Date', 'Merchant', 'Spend', 'Category', 'Account', 'Transaction ID'],
    table_rows: serializeTransactionsInOrder_(limited, limited.length || normalizeLimit_(args.limit, 8)),
    transactions_by_category: serializeTransactionsByCategory_(filtered, MAX_GROUPED_CATEGORY_EXAMPLES)
  };
}

function buildToolScope_(records, model, args) {
  const months = args.month ? extractMentionedMonths_(model, normalizeQueryText_(args.month)) : [];
  return {
    month: args.month || null,
    resolved_months: months,
    category: args.category || null,
    merchant: args.merchant || null,
    account: args.account || null,
    matched_transactions: records.length
  };
}

function filterRecordsByToolArgs_(records, model, args) {
  const monthMatches = args.month ? extractMentionedMonths_(model, normalizeQueryText_(args.month)) : [];
  const categoryQuery = normalizeQueryText_(args.category);
  const merchantQuery = normalizeQueryText_(args.merchant);
  const accountQuery = normalizeQueryText_(args.account);
  const expensesOnly = args.expensesOnly === true;
  const rules = model && model.activeRules ? model.activeRules : [];

  return records.filter(function(record) {
    if (!record.date) {
      return false;
    }
    if (expensesOnly && Number(record.amount || 0) >= 0) {
      return false;
    }
    if (isRecordExcludedByAnalyticsRules_(record, rules)) {
      return false;
    }
    if (monthMatches.length) {
      const monthKey = Utilities.formatDate(record.date, getSpreadsheetTimeZone_(), 'yyyy-MM');
      if (monthMatches.indexOf(monthKey) === -1) {
        return false;
      }
    }
    if (categoryQuery) {
      const detailed = normalizeQueryText_(formatDetailedCategoryLabel_(record.category));
      const family = normalizeQueryText_(formatCategoryLabel_(record.category));
      if (detailed.indexOf(categoryQuery) === -1 && family.indexOf(categoryQuery) === -1 && categoryQuery.indexOf(detailed) === -1) {
        return false;
      }
    }
    if (merchantQuery) {
      const merchant = normalizeQueryText_(record.name);
      if (merchant.indexOf(merchantQuery) === -1 && merchantQuery.indexOf(merchant) === -1) {
        return false;
      }
    }
    if (accountQuery) {
      const account = normalizeQueryText_(formatAccountLabel_(record.account) + ' ' + String(record.account || ''));
      if (account.indexOf(accountQuery) === -1 && accountQuery.indexOf(account) === -1) {
        return false;
      }
    }
    return true;
  });
}

function isRecordExcludedByAnalyticsRules_(record, rules) {
  if (!rules || !rules.length) {
    return false;
  }

  const labels = {
    account: formatAccountLabel_(record.account),
    category: formatCategoryLabel_(record.category),
    detailedCategory: formatDetailedCategoryLabel_(record.category),
    merchant: truncateLabel_(record.name, 36)
  };
  const decision = classifyRecordByRules_(record, rules, labels);
  return decision.treatment === 'exclude_from_analytics' || decision.treatment === 'exclude_from_cashflow';
}

function serializeNamedTotals_(list, limit) {
  return (list || []).slice(0, normalizeLimit_(limit, 6)).map(function(item) {
    return {
      name: item.name,
      total: roundCurrency_(item.total)
    };
  });
}

function serializeMerchantItems_(list, limit) {
  return (list || []).slice(0, normalizeLimit_(limit, 6)).map(function(item) {
    return {
      name: item.name,
      total: roundCurrency_(item.total),
      count: item.count || null,
      last_seen: item.lastSeen || null
    };
  });
}

function serializeTransactions_(records, limit) {
  return (records || [])
    .filter(function(record) {
      return record.date;
    })
    .sort(function(a, b) {
      return Math.abs(Number(b.amount || 0)) - Math.abs(Number(a.amount || 0));
    })
    .slice(0, normalizeLimit_(limit, 6))
    .map(function(record) {
      return serializeSingleTransaction_(record);
    });
}

function serializeTransactionsInOrder_(records, limit) {
  return (records || [])
    .filter(function(record) {
      return record.date;
    })
    .slice(0, normalizeLimit_(limit, 6))
    .map(serializeSingleTransaction_);
}

function serializeSingleTransaction_(record) {
  return {
    id: record.id,
    date: formatDateForPrompt_(record.date),
    name: truncateLabel_(record.name, 40),
    amount: roundCurrency_(record.amount),
    spend: roundCurrency_(Math.abs(Number(record.amount || 0))),
    category: formatDetailedCategoryLabel_(record.category),
    account: formatAccountLabel_(record.account)
  };
}

function serializeTransactionsByCategory_(records, examplesPerCategory) {
  const grouped = groupTransactionsByPrimaryCategory_(records);
  return Object.keys(grouped).map(function(category) {
    return {
      category: category,
      transactions: serializeTransactions_(grouped[category], examplesPerCategory || MAX_GROUPED_CATEGORY_EXAMPLES)
    };
  });
}

function normalizeLimit_(value, fallback) {
  const parsed = Number(value || fallback);
  if (!parsed || parsed < 1) {
    return fallback;
  }
  return Math.min(MAX_TOOL_TRANSACTION_RESULTS, Math.round(parsed));
}

function normalizeFunctionArgs_(args) {
  if (!args) {
    return {};
  }
  if (typeof args === 'string') {
    try {
      return JSON.parse(args);
    } catch (e) {
      return {};
    }
  }
  return args;
}

function extractMentionedCategories_(model, normalizedQuery) {
  return model.detailedCategoryList
    .map(function(item) {
      return item.name;
    })
    .filter(function(name) {
      const normalizedName = normalizeQueryText_(name);
      return normalizedName && normalizedQuery.indexOf(normalizedName) !== -1;
    })
    .slice(0, 5);
}

function extractMentionedMerchants_(model, normalizedQuery) {
  return model.merchantList
    .map(function(item) {
      return item.name;
    })
    .filter(function(name) {
      const normalizedName = normalizeQueryText_(name);
      return normalizedName.length >= 4 && normalizedQuery.indexOf(normalizedName) !== -1;
    })
    .slice(0, 5);
}

function extractMentionedAccounts_(model, normalizedQuery) {
  return model.accountList
    .map(function(item) {
      return item.name;
    })
    .filter(function(name) {
      const normalizedName = normalizeQueryText_(name);
      return normalizedName.length >= 4 && normalizedQuery.indexOf(normalizedName) !== -1;
    })
    .slice(0, 5);
}

function filterTransactionsByRetrieval_(records, filters) {
  return records.filter(function(record) {
    if (!record.date) {
      return false;
    }

    if (filters.months.length) {
      const monthKey = Utilities.formatDate(record.date, getSpreadsheetTimeZone_(), 'yyyy-MM');
      if (filters.months.indexOf(monthKey) === -1) {
        return false;
      }
    }

    if (filters.categories.length) {
      const detailedCategory = formatDetailedCategoryLabel_(record.category);
      if (filters.categories.indexOf(detailedCategory) === -1) {
        return false;
      }
    }

    if (filters.merchants.length) {
      const merchant = truncateLabel_(record.name, 36);
      if (filters.merchants.indexOf(merchant) === -1) {
        return false;
      }
    }

    if (filters.accounts && filters.accounts.length) {
      const account = formatAccountLabel_(record.account);
      if (filters.accounts.indexOf(account) === -1) {
        return false;
      }
    }

    return true;
  });
}

function selectRelevantTransactions_(records, query, intent, filters, evidenceLimit) {
  const normalized = String(query || '').toLowerCase();
  const keywords = normalized.split(/[^a-z0-9]+/).filter(function(token) {
    return token.length >= 4 && !isStopWord_(token);
  });
  const scopedRecords = filterTransactionsByRetrieval_(records, filters || { months: [], categories: [], merchants: [] });
  const sourceRecords = scopedRecords.length ? scopedRecords : records;

  const scored = sourceRecords.map(function(record) {
    const haystack = [
      String(record.name || '').toLowerCase(),
      String(record.category || '').toLowerCase(),
      String(record.account || '').toLowerCase(),
      formatDateForPrompt_(record.date).toLowerCase()
    ].join(' ');
    let score = 0;

    keywords.forEach(function(keyword) {
      if (haystack.indexOf(keyword) !== -1) {
        score += 3;
      }
    });

    if (intent.needsWeekend && record.date && isWeekendDate_(record.date)) {
      score += 2;
    }
    if (intent.needsAnomalies && Number(record.amount || 0) < 0) {
      score += Math.min(4, Math.round(Math.abs(record.amount) / 100));
    }
    if (intent.needsAdvice && Number(record.amount || 0) < 0) {
      score += 1;
    }

    return {
      record: record,
      score: score
    };
  }).filter(function(item) {
    return item.record.date;
  });

  const matched = scored.filter(function(item) {
    return item.score > 0;
  }).sort(compareScoredTransactions_);

  const fallback = scored.filter(function(item) {
    return Number(item.record.amount || 0) < 0;
  }).sort(compareScoredTransactions_);

  const selected = (matched.length ? matched : fallback).slice(0, evidenceLimit || MAX_EVIDENCE_TRANSACTIONS).map(function(item) {
    const spend = Number(item.record.amount || 0) < 0 ? formatCurrency_(Math.abs(item.record.amount)) : formatCurrency_(item.record.amount);
    return formatDateForPrompt_(item.record.date) + ' -> ' +
      truncateLabel_(item.record.name, 28) + ' -> ' +
      formatDetailedCategoryLabel_(item.record.category) + ' -> ' +
      truncateLabel_(formatAccountLabel_(item.record.account), 22) + ' -> ' +
      spend + ' -> (' + item.record.id + ')';
  });

  return selected.length ? selected : ['No relevant transactions found.'];
}

function compareScoredTransactions_(a, b) {
  if (b.score !== a.score) {
    return b.score - a.score;
  }
  const amountA = Math.abs(Number(a.record.amount || 0));
  const amountB = Math.abs(Number(b.record.amount || 0));
  if (amountB !== amountA) {
    return amountB - amountA;
  }
  return (b.record.date ? b.record.date.getTime() : 0) - (a.record.date ? a.record.date.getTime() : 0);
}

function normalizeQueryText_(text) {
  return String(text || '')
    .toLowerCase()
    .replace(/[-/]/g, ' ')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function isStopWord_(token) {
  return [
    'what', 'with', 'from', 'that', 'this', 'have', 'past', 'each', 'month',
    'months', 'include', 'breakdown', 'examples', 'example', 'analysis',
    'about', 'your', 'my', 'show', 'give', 'need', 'into', 'than', 'where',
    'when', 'which', 'should', 'could', 'would', 'their', 'them'
  ].indexOf(token) !== -1;
}

function parseAiIntent_(query, filters) {
  const normalized = String(query || '').toLowerCase();
  const needsMonthly = /(month|monthly|per month|each month|history|trend|past month)/.test(normalized) ||
    (filters && filters.months && filters.months.length > 0);
  const needsWeekly = /(weekly|calendar week|week over week|week-by-week)/.test(normalized);
  const needsWeekend = /(weekend|weekday|leak|leaks|day of week)/.test(normalized);
  const needsAccount = /(account|accounts|card|cards|bank|banks)/.test(normalized);
  const needsBreakdown = /(breakdown|broken down|split out|itemi[sz]e|by category|by account)/.test(normalized);
  const needsTransactions = /(transaction|transactions|example transaction|example transactions|show the rows|show rows|sample rows|transaction ids?)/.test(normalized);
  const needsGroupedTransactions = /(transaction|transactions|examples?|grouped|sample rows|merchant examples|transaction ids?|show the rows)/.test(normalized);
  const needsTabularOutput = /(table|tabular|markdown table|grid)/.test(normalized);
  const needsCategoryExamples = /((category|categories).*(breakdown|detail|details|example|examples|merchant|transaction id|transaction ids|sample|show me|include))|((breakdown|detail|details|example|examples|merchant|transaction id|transaction ids|sample|show me|include).*(category|categories))/.test(normalized);
  const needsAnomalies = /(largest|anomal|odd|unusual|biggest|outlier)/.test(normalized);
  const needsAdvice = /(optimi[sz]e|recommend|advice|should|plan|improve|cut|reduce|save money|save more|how can i|what should i do|help me)/.test(normalized);
  const needsCashflow = /(savings rate|save rate|real savings|cashflow|cash flow|net cashflow|income|external income|paycheck|payroll|burn rate|daily burn)/.test(normalized);
  const needsSpendFocus = /(spend|spending|expense|expenses|budget|save|cut|reduce|leak|leaks)/.test(normalized) ||
    needsMonthly ||
    needsWeekend ||
    needsCategoryExamples ||
    needsBreakdown ||
    needsCashflow;
  const needsGroundedEvidence = needsWeekend || needsBreakdown || needsTransactions || needsGroupedTransactions || needsTabularOutput || needsCategoryExamples || needsCashflow;
  return {
    needsMonthly: needsMonthly,
    needsWeekly: needsWeekly,
    needsWeekend: needsWeekend,
    needsAccount: needsAccount,
    needsSpendFocus: needsSpendFocus,
    needsBreakdown: needsBreakdown,
    needsTransactions: needsTransactions,
    needsGroupedTransactions: needsGroupedTransactions,
    needsTabularOutput: needsTabularOutput,
    needsCategoryExamples: needsCategoryExamples,
    needsAnomalies: needsAnomalies,
    needsAdvice: needsAdvice,
    needsCashflow: needsCashflow,
    needsGroundedEvidence: needsGroundedEvidence,
    needsStructuredReport: needsMonthly || needsWeekend || needsCategoryExamples || needsAnomalies || needsBreakdown || needsTabularOutput || needsCashflow
  };
}

function ensureGeminiKeyStatus_() {
  const props = getGeminiPropertyStores_();
  const sheetValue = sanitizeSettingValue_(getSetting_(GEMINI_SETTING_KEY));
  if (sheetValue) {
    props.script.setProperty(GEMINI_SETTING_KEY, sheetValue);
    props.user.setProperty(GEMINI_SETTING_KEY, sheetValue);
    setSetting_(GEMINI_SETTING_KEY, GEMINI_KEY_MIGRATED_MARKER);
    return;
  }

  const storedValue = sanitizeSettingValue_(props.script.getProperty(GEMINI_SETTING_KEY)) ||
    sanitizeSettingValue_(props.user.getProperty(GEMINI_SETTING_KEY));
  if (storedValue) {
    setSetting_(GEMINI_SETTING_KEY, GEMINI_KEY_MIGRATED_MARKER);
  }
}

function resetGeminiKeyStorage() {
  clearGeminiKeyStorage_();
  SpreadsheetApp.getUi().alert('Gemini key storage cleared. Paste your current Gemini API key into Settings!B2, save the cell, then open the sidebar or refresh visuals.');
}

function clearGeminiKeyStorage_() {
  const props = getGeminiPropertyStores_();
  props.script.deleteProperty(GEMINI_SETTING_KEY);
  props.user.deleteProperty(GEMINI_SETTING_KEY);
  setSetting_(GEMINI_SETTING_KEY, '');
}

function getGeminiApiKey_() {
  const props = getGeminiPropertyStores_();
  const scriptValue = sanitizeSettingValue_(props.script.getProperty(GEMINI_SETTING_KEY));
  const userValue = sanitizeSettingValue_(props.user.getProperty(GEMINI_SETTING_KEY));
  const sheetValue = sanitizeSettingValue_(getSetting_(GEMINI_SETTING_KEY));
  const storedValue = scriptValue || userValue;

  if (sheetValue && sheetValue !== storedValue) {
    props.script.setProperty(GEMINI_SETTING_KEY, sheetValue);
    props.user.setProperty(GEMINI_SETTING_KEY, sheetValue);
    setSetting_(GEMINI_SETTING_KEY, GEMINI_KEY_MIGRATED_MARKER);
    return sheetValue;
  }

  if (storedValue) {
    setSetting_(GEMINI_SETTING_KEY, GEMINI_KEY_MIGRATED_MARKER);
    return storedValue;
  }

  return '';
}

function buildGeminiRequest_(systemInstruction, userPrompt) {
  return {
    systemInstruction: {
      role: 'system',
      parts: [{ text: systemInstruction }]
    },
    contents: [{
      role: 'user',
      parts: [{ text: userPrompt }]
    }],
    generationConfig: {
      temperature: 0.2,
      topP: 0.9,
      maxOutputTokens: 1400
    }
  };
}

function _callGemini(generateRequest, apiKey, tokenEstimate) {
  const payload = _callGeminiPayload_(generateRequest, apiKey, tokenEstimate);
  const candidates = payload.candidates || [];
  const parts = candidates.length && candidates[0].content && candidates[0].content.parts
    ? candidates[0].content.parts
    : [];
  const text = parts.map(function(part) {
    return part.text || '';
  }).join('').trim();

  if (!text) {
    const finishReason = candidates.length ? candidates[0].finishReason : null;
    throw new Error(finishReason ? 'Gemini returned no text (' + finishReason + ')' : 'Gemini returned no text.');
  }

  return text;
}

function _callGeminiPayload_(generateRequest, apiKey, tokenEstimate) {
  let lastError = null;
  for (let index = 0; index < GEMINI_MODEL_CHAIN.length; index++) {
    const modelName = GEMINI_MODEL_CHAIN[index];
    const response = UrlFetchApp.fetch(
      'https://generativelanguage.googleapis.com/v1beta/models/' + modelName + ':generateContent?key=' + encodeURIComponent(apiKey),
      {
        method: 'post',
        contentType: 'application/json',
        muteHttpExceptions: true,
        payload: JSON.stringify(generateRequest)
      }
    );

    const status = response.getResponseCode();
    const payload = JSON.parse(response.getContentText() || '{}');
    if (status >= 400) {
      const message = payload.error && payload.error.message ? payload.error.message : 'Gemini request failed with status ' + status;
      if (shouldRetryGeminiWithNextModel_(status, message, index)) {
        lastError = new Error(message);
        continue;
      }
      throw new Error(message);
    }

    if (payload.promptFeedback && payload.promptFeedback.blockReason) {
      throw new Error('Prompt blocked: ' + payload.promptFeedback.blockReason);
    }

    payload._modelUsed = modelName;
    storeLastAiUsage_(tokenEstimate, payload.usageMetadata || {}, modelName);
    return payload;
  }

  throw lastError || new Error('Gemini request failed across all configured models.');
}

function _countGeminiTokens(generateRequest, apiKey) {
  try {
    for (let index = 0; index < GEMINI_MODEL_CHAIN.length; index++) {
      const modelName = GEMINI_MODEL_CHAIN[index];
      const response = UrlFetchApp.fetch(
        'https://generativelanguage.googleapis.com/v1beta/models/' + modelName + ':countTokens?key=' + encodeURIComponent(apiKey),
        {
          method: 'post',
          contentType: 'application/json',
          muteHttpExceptions: true,
          payload: JSON.stringify({
            generateContentRequest: generateRequest
          })
        }
      );
      if (response.getResponseCode() >= 400) {
        continue;
      }
      const payload = JSON.parse(response.getContentText() || '{}');
      return {
        totalTokens: Number(payload.totalTokens || 0),
        cachedContentTokenCount: Number(payload.cachedContentTokenCount || 0)
      };
    }
    return null;
  } catch (e) {
    return null;
  }
}

function shouldRetryGeminiWithNextModel_(status, message, index) {
  if (index >= GEMINI_MODEL_CHAIN.length - 1) {
    return false;
  }
  const normalized = String(message || '').toLowerCase();
  return status === 429 ||
    status === 500 ||
    status === 503 ||
    normalized.indexOf('quota') !== -1 ||
    normalized.indexOf('rate limit') !== -1 ||
    normalized.indexOf('retry') !== -1 ||
    normalized.indexOf('resource exhausted') !== -1;
}

function storeLastAiUsage_(tokenEstimate, usageMetadata, modelName) {
  const payload = {
    promptTokens: Number((usageMetadata && usageMetadata.promptTokenCount) || (tokenEstimate && tokenEstimate.totalTokens) || 0),
    outputTokens: Number((usageMetadata && usageMetadata.candidatesTokenCount) || 0),
    totalTokens: Number((usageMetadata && usageMetadata.totalTokenCount) || ((tokenEstimate && tokenEstimate.totalTokens) || 0)),
    model: modelName || ''
  };
  PropertiesService.getUserProperties().setProperty(LAST_AI_USAGE_KEY, JSON.stringify(payload));
}

function getLastAiUsage_() {
  try {
    const raw = PropertiesService.getUserProperties().getProperty(LAST_AI_USAGE_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null;
  }
}

function getGeminiPropertyStores_() {
  return {
    script: PropertiesService.getScriptProperties(),
    user: PropertiesService.getUserProperties()
  };
}

function sanitizeSettingValue_(value) {
  const text = (value || '').toString().trim();
  if (!text || text === GEMINI_KEY_MIGRATED_MARKER || text === LEGACY_GEMINI_KEY_MIGRATED_MARKER) {
    return '';
  }
  return text;
}

function getSetting_(key) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  if (!sheet) {
    return null;
  }
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === key) {
      return data[i][1];
    }
  }
  return null;
}

function setSetting_(key, value) {
  const sheet = ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), 'Settings');
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 2, 2).setValue(value);
      return;
    }
  }

  sheet.appendRow([key, value]);
}

function setupRulesSheet_(sheet) {
  if (!sheet) {
    return;
  }

  seedRulesSheet_(sheet);
  writeRulesLegend_(sheet);
}

function seedRulesSheet_(sheet) {
  if (!sheet || sheet.getLastRow() > 1) {
    return;
  }

  sheet.getRange(2, 1, 4, 7).setValues([
    ['rule_001', false, 'account', 'exact', 'Example Business Card ending 0000', 'exclude_from_analytics', 'Enable and edit to hide an account from Dashboard/Insights/AI analytics.'],
    ['rule_002', false, 'category', 'contains', 'Internal Transfer', 'exclude_from_cashflow', 'Enable if Plaid categorizes a transfer pattern that hardcoded internal-transfer detection misses.'],
    ['rule_003', false, 'merchant', 'contains', 'VENMO', 'review_only', 'Enable to keep spend included but flag mixed/reimbursement merchants for AI context.'],
    ['rule_004', false, 'category', 'contains', 'Loan Payments Credit Card Payment', 'exclude_from_cashflow', 'Enable if a credit-card payment category is not already excluded automatically.']
  ]);
}

function writeRulesLegend_(sheet) {
  sheet.getRange('I1:K14').clearContent().clearFormat();
  sheet.getRange('I1:K1').merge()
    .setValue('Rules Legend')
    .setBackground('#0f766e')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.getRange('I2:K14').setValues([
    ['Field', 'Allowed Values', 'Meaning'],
    ['Enabled', 'TRUE', 'Rule is active'],
    ['Enabled', 'FALSE', 'Rule is ignored'],
    ['Rule Type', 'account', 'Match the Account column / friendly card name'],
    ['Rule Type', 'category', 'Match category or detailed category'],
    ['Rule Type', 'merchant', 'Match transaction merchant/name'],
    ['Match Type', 'exact', 'Must match the whole normalized value'],
    ['Match Type', 'contains', 'Match Value may appear anywhere in the target'],
    ['Treatment', 'include', 'Keep included; lowest priority if other rules match'],
    ['Treatment', 'exclude_from_analytics', 'Remove from Dashboard, Insights, Analytics, and grounded AI totals'],
    ['Treatment', 'exclude_from_cashflow', 'Remove from spend/income/cashflow totals but keep raw row'],
    ['Treatment', 'review_only', 'Keep included but flag for AI/review context'],
    ['Priority', 'exclude > cashflow > review > include', 'When multiple enabled rules match, strongest treatment wins']
  ]);
  sheet.getRange('I2:K2')
    .setBackground('#134e4a')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  sheet.getRange('I2:K14')
    .setBorder(true, true, true, true, true, true, '#94a3b8', SpreadsheetApp.BorderStyle.SOLID)
    .setWrap(true)
    .setVerticalAlignment('top');
  sheet.setColumnWidths(9, 3, 180);
}

function getAnalyticsRules_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Rules');
  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  return rows.map(function(row) {
    return {
      id: String(row[0] || '').trim(),
      enabled: isRuleEnabled_(row[1]),
      ruleType: String(row[2] || '').trim().toLowerCase(),
      matchType: String(row[3] || '').trim().toLowerCase(),
      matchValue: String(row[4] || '').trim(),
      treatment: String(row[5] || '').trim().toLowerCase(),
      notes: String(row[6] || '').trim()
    };
  }).filter(function(rule) {
    return rule.enabled &&
      ['account', 'category', 'merchant'].indexOf(rule.ruleType) !== -1 &&
      ['exact', 'contains'].indexOf(rule.matchType) !== -1 &&
      ['include', 'exclude_from_analytics', 'exclude_from_cashflow', 'review_only'].indexOf(rule.treatment) !== -1 &&
      rule.matchValue;
  });
}

function isRuleEnabled_(value) {
  if (value === true) {
    return true;
  }
  const normalized = String(value || '').trim().toLowerCase();
  return ['true', 'yes', 'y', '1', 'enabled'].indexOf(normalized) !== -1;
}

function classifyRecordByRules_(record, rules, labels) {
  const precedence = {
    include: 0,
    review_only: 1,
    exclude_from_cashflow: 2,
    exclude_from_analytics: 3
  };
  let selected = { treatment: 'include', precedence: 0, matchedRules: [] };

  (rules || []).forEach(function(rule) {
    if (!doesRuleMatchRecord_(rule, record, labels)) {
      return;
    }

    const rank = precedence[rule.treatment] || 0;
    selected.matchedRules.push(rule.id || rule.matchValue);
    if (rank > selected.precedence) {
      selected = {
        treatment: rule.treatment,
        precedence: rank,
        matchedRules: selected.matchedRules
      };
    }
  });

  return selected;
}

function doesRuleMatchRecord_(rule, record, labels) {
  const value = normalizeRuleComparable_(rule.matchValue);
  if (!value) {
    return false;
  }

  const targets = getRuleTargets_(rule.ruleType, record, labels).map(normalizeRuleComparable_);
  if (rule.matchType === 'exact') {
    return targets.some(function(target) {
      return target === value;
    });
  }

  return targets.some(function(target) {
    return target.indexOf(value) !== -1;
  });
}

function getRuleTargets_(ruleType, record, labels) {
  if (ruleType === 'account') {
    return [labels.account, record.account];
  }
  if (ruleType === 'category') {
    return [labels.category, labels.detailedCategory, record.category];
  }
  if (ruleType === 'merchant') {
    return [labels.merchant, record.name];
  }
  return [];
}

function normalizeRuleComparable_(value) {
  return String(value || '').toLowerCase().replace(/\s+/g, ' ').trim();
}

function summarizeAnalyticsRules_(rules) {
  if (!rules || !rules.length) {
    return 'None';
  }
  return rules.slice(0, 8).map(function(rule) {
    return rule.ruleType + ' ' + rule.matchType + ' "' + rule.matchValue + '" -> ' + rule.treatment;
  }).join('; ');
}

function getTransactionRecords_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues()
    .filter(function(row) {
      return row[0];
    })
    .map(function(row) {
      return {
        id: String(row[0]),
        date: coerceSheetDate_(row[1]),
        name: String(row[2] || 'Unknown Merchant'),
        amount: Number(row[3] || 0),
        category: String(row[4] || 'Uncategorized'),
        account: String(row[5] || 'Unknown Account'),
        pending: String(row[6]).toLowerCase() === 'true' || row[6] === true
      };
    });
}

function coerceSheetDate_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }
  if (typeof value === 'number') {
    return new Date(Math.round((value - 25569) * 86400 * 1000));
  }
  const parsed = new Date(value);
  if (isNaN(parsed.getTime())) {
    return null;
  }
  return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
}

function getWeekStart_(date) {
  const value = new Date(date.getTime());
  const day = value.getDay();
  const diff = day === 0 ? -6 : 1 - day;
  value.setDate(value.getDate() + diff);
  value.setHours(0, 0, 0, 0);
  return value;
}

function isWeekendDate_(date) {
  const day = date.getDay();
  return day === 0 || day === 6;
}

function createCategoryStat_(name) {
  return { name: name, total: 0, examples: [] };
}

function createMerchantStat_(name) {
  return { name: name, total: 0, count: 0, lastSeen: null };
}

function createExample_(record, amount, timezone) {
  return {
    id: record.id,
    name: truncateLabel_(record.name, 32),
    amount: amount,
    date: Utilities.formatDate(record.date, timezone, 'yyyy-MM-dd')
  };
}

function buildScopedMonthDetailLines_(records, model, monthKeys) {
  const lines = [];
  (monthKeys || []).forEach(function(monthKey) {
    const bucket = model.months[monthKey];
    if (!bucket) {
      return;
    }
    const monthRecords = filterRecordsByToolArgs_(records, model, {
      month: monthKey,
      expensesOnly: true
    });
    lines.push('Month -> ' + monthKey + ' (' + formatMonthDisplayLabel_(monthKey) + ')');
    lines.push('Full Account Breakdown -> [' + formatBucketSummary_(bucket.accountList, Math.max(bucket.accountList.length, MAX_TOP_ITEMS)) + ']');
    lines.push('Full Category Breakdown -> [' + formatBucketSummary_(bucket.categoryList, Math.max(bucket.categoryList.length, MAX_TOP_ITEMS)) + ']');
    lines.push('Grouped Transactions By Category ->');
    buildGroupedCategoryExampleLines_(monthRecords).forEach(function(line) {
      lines.push(line);
    });
    lines.push('Examples -> [' + buildExampleList_(bucket.examples, MAX_CATEGORY_EXAMPLES) + ']');
  });
  return lines.length ? lines : ['No scoped month detail available.'];
}

function buildGroupedCategoryExampleLines_(records) {
  const grouped = groupTransactionsByPrimaryCategory_(records);
  const categories = Object.keys(grouped).sort(function(a, b) {
    return grouped[b].reduce(sumSpend_, 0) - grouped[a].reduce(sumSpend_, 0);
  });
  if (!categories.length) {
    return ['No category-grouped transactions found.'];
  }
  return categories.map(function(category) {
    const examples = grouped[category]
      .sort(compareTransactionSpendDesc_)
      .slice(0, MAX_GROUPED_CATEGORY_EXAMPLES)
      .map(function(record) {
        return truncateLabel_(record.name, 28) + ' ' + formatCurrency_(Math.abs(Number(record.amount || 0))) + ' (' + record.id + ')';
      })
      .join(', ');
    return category + ' -> [' + examples + ']';
  });
}

function groupTransactionsByPrimaryCategory_(records) {
  return (records || []).reduce(function(groups, record) {
    if (!record || Number(record.amount || 0) >= 0) {
      return groups;
    }
    const key = formatCategoryLabel_(record.category);
    groups[key] = groups[key] || [];
    groups[key].push(record);
    return groups;
  }, {});
}

function compareTransactionSpendDesc_(a, b) {
  return Math.abs(Number(b.amount || 0)) - Math.abs(Number(a.amount || 0));
}

function sumSpend_(sum, record) {
  return sum + Math.abs(Number(record.amount || 0));
}

function appendAdviceSections_(lines, model) {
  const sections = buildAdviceSections_(model);
  if (!sections.quickWins.length && !sections.subscriptions.length && !sections.behaviorPatterns.length && !sections.watchList.length) {
    return;
  }
  lines.push('');
  lines.push('Quick Wins:');
  (sections.quickWins.length ? sections.quickWins : ['No obvious quick wins identified from the current scope.']).forEach(function(item) {
    lines.push('- ' + item);
  });
  lines.push('');
  lines.push('Subscriptions:');
  (sections.subscriptions.length ? sections.subscriptions : ['No recurring subscription-like merchants stand out in the current scope.']).forEach(function(item) {
    lines.push('- ' + item);
  });
  lines.push('');
  lines.push('Behavior Patterns:');
  (sections.behaviorPatterns.length ? sections.behaviorPatterns : ['No dominant behavior pattern stood out in the current scope.']).forEach(function(item) {
    lines.push('- ' + item);
  });
  lines.push('');
  lines.push('Watch List:');
  (sections.watchList.length ? sections.watchList : ['No immediate watch-list items beyond normal bills.']).forEach(function(item) {
    lines.push('- ' + item);
  });
}

function buildAdviceSections_(model) {
  const quickWins = [];
  const subscriptions = [];
  const behaviorPatterns = [];
  const watchList = [];
  const discretionaryCategories = ['Food & Drink', 'Shopping', 'Services', 'Travel', 'Entertainment', 'Transportation'];

  model.categoryList.forEach(function(item) {
    if (quickWins.length >= 3) {
      return;
    }
    if (discretionaryCategories.indexOf(item.name) !== -1 && item.total > 0) {
      quickWins.push(item.name + ' is running at ' + formatCurrency_(item.total) + '; cutting even 10-15% here would move the needle fastest.');
    }
  });

  model.recurringCandidates.slice(0, 3).forEach(function(item) {
    subscriptions.push(item.name + ' appears ' + item.count + ' times for ' + formatCurrency_(item.total) + ' total. Confirm it is still worth keeping.');
  });

  if (model.totalSpend > 0) {
    const weekendShare = model.weekendSpend / model.totalSpend;
    behaviorPatterns.push('Weekend spend is ' + formatPercent_(weekendShare) + ' of total external spend.');
  }
  if (model.merchantList.length) {
    behaviorPatterns.push('Top merchant concentration is ' + model.merchantList[0].name + ' at ' + formatCurrency_(model.merchantList[0].total) + '.');
  }

  model.anomalies.slice(0, 3).forEach(function(item) {
    watchList.push(item.merchant + ' on ' + item.date + ' for ' + formatCurrency_(item.spend) + ' is worth confirming as intentional.');
  });

  return {
    quickWins: quickWins,
    subscriptions: subscriptions,
    behaviorPatterns: behaviorPatterns,
    watchList: watchList
  };
}

function getOrCreateMonthBucket_(months, key) {
  if (!months[key]) {
    months[key] = {
      key: key,
      income: 0,
      spend: 0,
      net: 0,
      transactionCount: 0,
      weekendSpend: 0,
      weekdaySpend: 0,
      excludedOutflow: 0,
      excludedInflow: 0,
      excludedCount: 0,
      ruleExcludedOutflow: 0,
      ruleExcludedInflow: 0,
      ruleExcludedCount: 0,
      reviewOnlyCount: 0,
      accounts: {},
      categories: {},
      merchants: {},
      examples: []
    };
  }
  return months[key];
}

function getOrCreateWeekBucket_(weeks, key) {
  if (!weeks[key]) {
    weeks[key] = {
      key: key,
      income: 0,
      spend: 0,
      net: 0,
      transactionCount: 0
    };
  }
  return weeks[key];
}

function initializeWeekdayTotals_() {
  const totals = {};
  WEEKDAY_ORDER.forEach(function(day) {
    totals[day] = 0;
  });
  return totals;
}

function createSection_(row, col, values) {
  return {
    row: row,
    col: col,
    values: values
  };
}

function getSectionRange_(sheet, section, widthOverrideStart, widthOverrideEnd) {
  const colOffset = widthOverrideStart ? widthOverrideStart - 1 : 0;
  const width = widthOverrideEnd ? widthOverrideEnd - colOffset : section.values[0].length;
  return sheet.getRange(section.row, section.col + colOffset, section.values.length, width);
}

function writeTable_(sheet, row, col, values, headerColor) {
  const width = values[0].length;
  const range = sheet.getRange(row, col, values.length, width);
  range.setValues(values);
  range.offset(0, 0, 1, width)
    .setBackground(headerColor || '#1f2937')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  if (values.length > 1) {
    range.offset(1, 0, values.length - 1, width).setBackground('#ffffff');
  }
}

function writeVisibleSectionHeader_(sheet, row, col, label, width, color) {
  sheet.getRange(row, col, 1, width || 4).merge()
    .setValue(label)
    .setBackground(color || '#1f2937')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('left');
}

function clearCharts_(sheet) {
  sheet.getCharts().forEach(function(chart) {
    sheet.removeChart(chart);
  });
}

function insertCharts_(sheet, charts) {
  charts.forEach(function(chart) {
    sheet.insertChart(chart);
  });
}

function ensureSheet_(spreadsheet, name) {
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
  }
  return sheet;
}

function bucketMapToSortedList_(bucketMap) {
  return Object.keys(bucketMap).map(function(name) {
    return { name: name, total: bucketMap[name] };
  }).sort(function(a, b) {
    return b.total - a.total;
  });
}

function categoryMapToSortedList_(categoryMap) {
  return Object.keys(categoryMap).map(function(name) {
    const entry = categoryMap[name];
    entry.examples = entry.examples.sort(function(a, b) {
      return b.amount - a.amount;
    }).slice(0, MAX_CATEGORY_EXAMPLES);
    return entry;
  }).sort(function(a, b) {
    return b.total - a.total;
  });
}

function merchantMapToSortedList_(merchantMap) {
  return Object.keys(merchantMap).map(function(name) {
    return merchantMap[name];
  }).sort(function(a, b) {
    if (b.total !== a.total) {
      return b.total - a.total;
    }
    return b.count - a.count;
  });
}

function selectDetailedCategories_(query, categories) {
  const normalized = String(query || '').toLowerCase();
  const matched = categories.filter(function(categoryInfo) {
    const tokens = categoryInfo.name.toLowerCase().split(/[^a-z0-9]+/).filter(function(token) {
      return token.length >= 4;
    });
    return tokens.some(function(token) {
      return normalized.indexOf(token) !== -1;
    });
  });

  if (matched.length) {
    return matched.slice(0, MAX_TOP_ITEMS);
  }
  return categories.slice(0, MAX_TOP_ITEMS);
}

function formatBucketSummary_(items, limit) {
  if (!items || items.length === 0) {
    return 'None';
  }
  return items.slice(0, limit).map(function(item) {
    return item.name + ' ' + formatCurrency_(item.total);
  }).join(', ');
}

function buildExampleList_(examples, limit) {
  if (!examples || examples.length === 0) {
    return 'No examples';
  }
  return examples.slice(0, limit).map(function(example) {
    return example.name + ' (' + example.id + ')';
  }).join(', ');
}

function formatAccountLabel_(accountId) {
  const text = String(accountId || 'Unknown Account');
  if (text.indexOf(' - ') !== -1 || /\bending\s+\d{2,4}$/i.test(text)) {
    return text;
  }
  if (text.length <= 12) {
    return text;
  }
  return 'Acct ...' + text.slice(-6);
}

function formatCategoryLabel_(category) {
  const raw = String(category || 'Uncategorized');
  const family = raw.indexOf('>') !== -1 ? raw.split('>')[0].trim() : raw;
  return humanizePlaidCategorySegment_(family);
}

function formatDetailedCategoryLabel_(category) {
  const raw = String(category || 'Uncategorized');
  return raw.split(/\s*>\s*/).map(function(part) {
    return humanizePlaidCategorySegment_(part);
  }).join(' / ');
}

function humanizePlaidCategorySegment_(segment) {
  const normalized = String(segment || 'Uncategorized').trim().toUpperCase();
  const aliasMap = {
    'FOOD_AND_DRINK': 'Food & Drink',
    'GOVERNMENT_AND_NON_PROFIT': 'Government & Nonprofit',
    'GENERAL_SERVICES': 'Services',
    'GENERAL_MERCHANDISE': 'Shopping',
    'RENT_AND_UTILITIES': 'Rent & Utilities',
    'LOAN_PAYMENTS': 'Debt Payments',
    'TRANSFER_OUT': 'Internal Transfer Out',
    'TRANSFER_IN': 'Internal Transfer In',
    'HOME_IMPROVEMENT': 'Home Improvement',
    'PERSONAL_CARE': 'Personal Care',
    'TRANSPORTATION': 'Transportation',
    'ENTERTAINMENT': 'Entertainment',
    'TRAVEL': 'Travel',
    'MEDICAL': 'Medical'
  };
  if (aliasMap[normalized]) {
    return aliasMap[normalized];
  }
  return toTitleCase_(normalized.replace(/_/g, ' ').toLowerCase())
    .replace(/\bAnd\b/g, '&')
    .replace(/\bTv\b/g, 'TV')
    .replace(/\bNon Profit\b/g, 'Nonprofit');
}

function classifyCashflowRecord_(record, categoryLabel, detailedCategoryLabel) {
  const normalizedName = normalizeQueryText_(record && record.name);
  const primary = String(categoryLabel || '');
  const detailed = String(detailedCategoryLabel || '');
  const isCreditCardPayment = detailed === 'Loan Payments / Loan Payments Credit Card Payment' ||
    /payment thank you|autopay payment|online ach payment|automatic payment|des ach pmt|des ccpymt|credit card payment/.test(normalizedName);
  const isAccountTransfer = primary === 'Transfer In' ||
    primary === 'Transfer Out' ||
    detailed === 'Transfer In / Transfer In Account Transfer' ||
    detailed === 'Transfer Out / Transfer Out Account Transfer';

  return {
    isCreditCardPayment: isCreditCardPayment,
    isAccountTransfer: isAccountTransfer,
    excludeFromCashflow: isCreditCardPayment || isAccountTransfer
  };
}

function truncateLabel_(text, limit) {
  const value = String(text || '');
  if (value.length <= limit) {
    return value;
  }
  return value.slice(0, limit - 1) + '…';
}

function toTitleCase_(text) {
  return String(text || '')
    .toLowerCase()
    .replace(/\b\w/g, function(char) { return char.toUpperCase(); });
}

function formatCurrency_(value) {
  return '$' + roundCurrency_(value).toFixed(2);
}

function formatPercent_(value) {
  return (Number(value || 0) * 100).toFixed(1) + '%';
}

function roundCurrency_(value) {
  return Math.round(Number(value || 0) * 100) / 100;
}

function formatDateForPrompt_(date) {
  if (!date) {
    return 'Unknown';
  }
  return Utilities.formatDate(date, getSpreadsheetTimeZone_(), 'yyyy-MM-dd');
}

function formatMonthDisplayLabel_(monthKey) {
  if (!monthKey || !/^\d{4}-\d{2}$/.test(String(monthKey))) {
    return String(monthKey || 'Unknown');
  }
  const parts = String(monthKey).split('-');
  const date = new Date(Number(parts[0]), Number(parts[1]) - 1, 1);
  return Utilities.formatDate(date, getSpreadsheetTimeZone_(), 'MMM yyyy');
}

function getSpreadsheetTimeZone_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || Session.getScriptTimeZone();
}

function uniqueValue_(value, index, array) {
  return array.indexOf(value) === index;
}
