// asktru.TaskZoom — script.js
// Smart Task Filtering for NotePlan

// ============================================
// CONFIGURATION
// ============================================

var PLUGIN_ID = 'asktru.TaskZoom';
var WINDOW_ID = 'asktru.TaskZoom.dashboard';

function getSettings() {
  const settings = DataStore.settings || {};
  const excludeStr = settings.foldersToExclude || '@Archive, @Trash, @Templates';
  let savedFilters = [];
  try {
    savedFilters = JSON.parse(settings.savedFilters || '[]');
  } catch (e) { savedFilters = []; }
  return {
    foldersToExclude: excludeStr.split(',').map(s => s.trim()).filter(Boolean),
    savedFilters: savedFilters,
    defaultGroupBy: settings.defaultGroupBy || 'note',
    showCompletedTasks: settings.showCompletedTasks === true || settings.showCompletedTasks === 'true',
  };
}

function saveFilters(filters) {
  const settings = DataStore.settings || {};
  settings.savedFilters = JSON.stringify(filters);
  DataStore.settings = settings;
}

function saveUserPrefs(filterId, query, groupBy) {
  const settings = DataStore.settings || {};
  settings.lastFilterId = filterId || '';
  settings.lastQuery = query || '';
  // Save groupBy per filter
  var groupByMap = {};
  try { groupByMap = JSON.parse(settings.filterGroupByMap || '{}'); } catch (e) { groupByMap = {}; }
  if (filterId) groupByMap[filterId] = groupBy || 'note';
  settings.filterGroupByMap = JSON.stringify(groupByMap);
  DataStore.settings = settings;
}

function getUserPrefs() {
  const settings = DataStore.settings || {};
  var groupByMap = {};
  try { groupByMap = JSON.parse(settings.filterGroupByMap || '{}'); } catch (e) { groupByMap = {}; }
  return {
    lastFilterId: settings.lastFilterId || '',
    lastQuery: settings.lastQuery || '',
    groupByMap: groupByMap,
  };
}

// ============================================
// DATE UTILITIES
// ============================================

function getTodayStr() {
  var d = new Date();
  return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
}

function getTomorrowStr() {
  var d = new Date(); d.setDate(d.getDate()+1);
  return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
}

function getYesterdayStr() {
  var d = new Date(); d.setDate(d.getDate()-1);
  return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
}

function getStartOfWeek() {
  var d = new Date();
  var day = d.getDay(); // 0=Sun
  var diff = d.getDate() - day + (day === 0 ? -6 : 1); // Monday
  var mon = new Date(d); mon.setDate(diff);
  return mon.getFullYear() + '-' + String(mon.getMonth()+1).padStart(2,'0') + '-' + String(mon.getDate()).padStart(2,'0');
}

function getEndOfWeek() {
  var d = new Date();
  var day = d.getDay();
  var diff = d.getDate() - day + (day === 0 ? 0 : 7); // Sunday
  var sun = new Date(d); sun.setDate(diff);
  return sun.getFullYear() + '-' + String(sun.getMonth()+1).padStart(2,'0') + '-' + String(sun.getDate()).padStart(2,'0');
}

function getStartOfNextWeek() {
  var d = new Date();
  var day = d.getDay();
  var diff = d.getDate() - day + (day === 0 ? 1 : 8); // Next Monday
  var mon = new Date(d); mon.setDate(diff);
  return mon.getFullYear() + '-' + String(mon.getMonth()+1).padStart(2,'0') + '-' + String(mon.getDate()).padStart(2,'0');
}

function getEndOfNextWeek() {
  var d = new Date();
  var day = d.getDay();
  var diff = d.getDate() - day + (day === 0 ? 7 : 14); // Next Sunday
  var sun = new Date(d); sun.setDate(diff);
  return sun.getFullYear() + '-' + String(sun.getMonth()+1).padStart(2,'0') + '-' + String(sun.getDate()).padStart(2,'0');
}

function getStartOfMonth() {
  var d = new Date();
  return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-01';
}

function getEndOfMonth() {
  var d = new Date();
  var last = new Date(d.getFullYear(), d.getMonth()+1, 0);
  return last.getFullYear() + '-' + String(last.getMonth()+1).padStart(2,'0') + '-' + String(last.getDate()).padStart(2,'0');
}

function getStartOfNextMonth() {
  var d = new Date();
  var nm = new Date(d.getFullYear(), d.getMonth()+1, 1);
  return nm.getFullYear() + '-' + String(nm.getMonth()+1).padStart(2,'0') + '-01';
}

function getEndOfNextMonth() {
  var d = new Date();
  var last = new Date(d.getFullYear(), d.getMonth()+2, 0);
  return last.getFullYear() + '-' + String(last.getMonth()+1).padStart(2,'0') + '-' + String(last.getDate()).padStart(2,'0');
}

function formatDate(dateStr) {
  if (!dateStr) return '';
  var d = new Date(dateStr + 'T00:00:00');
  var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var now = new Date();
  if (d.getFullYear() === now.getFullYear()) {
    return months[d.getMonth()] + ' ' + d.getDate();
  }
  return months[d.getMonth()] + ' ' + d.getDate() + ', ' + d.getFullYear();
}

function getISOWeek(dateStr) {
  var d = new Date(dateStr + 'T00:00:00');
  d.setHours(0,0,0,0);
  d.setDate(d.getDate() + 3 - (d.getDay() + 6) % 7);
  var week1 = new Date(d.getFullYear(), 0, 4);
  var weekNum = 1 + Math.round(((d - week1) / 86400000 - 3 + (week1.getDay() + 6) % 7) / 7);
  return d.getFullYear() + '-W' + String(weekNum).padStart(2, '0');
}

// ============================================
// TASK SCANNING
// ============================================

function encSafe(str) {
  return encodeURIComponent(str).replace(/[!'()*]/g, function(c) { return '%' + c.charCodeAt(0).toString(16).toUpperCase(); });
}

function decSafe(str) {
  try { return decodeURIComponent(str); } catch (e) { return str; }
}

/**
 * Determine if a note is a calendar note and extract its implicit date.
 * Calendar note filenames: 20260320.md (daily), 2026-W12.md (weekly),
 * 2026-03.md (monthly), 2026-Q1.md (quarterly), 2026.md (yearly)
 */
function getCalendarNoteInfo(note) {
  var filename = note.filename || '';
  var baseName = filename.replace(/\.(md|txt)$/, '');

  // Daily: YYYYMMDD
  var dailyMatch = baseName.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (dailyMatch) {
    return {
      isCalendar: true,
      calendarType: 'day',
      implicitDate: dailyMatch[1] + '-' + dailyMatch[2] + '-' + dailyMatch[3],
      implicitWeek: null,
      displayTitle: dailyMatch[1] + '-' + dailyMatch[2] + '-' + dailyMatch[3],
    };
  }

  // Weekly: YYYY-Www
  var weeklyMatch = baseName.match(/^(\d{4}-W\d{2})$/);
  if (weeklyMatch) {
    return {
      isCalendar: true,
      calendarType: 'week',
      implicitDate: null,
      implicitWeek: weeklyMatch[1],
      displayTitle: weeklyMatch[1],
    };
  }

  // Monthly: YYYY-MM
  var monthlyMatch = baseName.match(/^(\d{4})-(\d{2})$/);
  if (monthlyMatch) {
    return {
      isCalendar: true,
      calendarType: 'month',
      implicitDate: monthlyMatch[1] + '-' + monthlyMatch[2] + '-01',
      implicitWeek: null,
      displayTitle: monthlyMatch[1] + '-' + monthlyMatch[2],
    };
  }

  // Quarterly: YYYY-Qn
  var quarterlyMatch = baseName.match(/^(\d{4})-Q([1-4])$/);
  if (quarterlyMatch) {
    var qMonth = String((parseInt(quarterlyMatch[2], 10) - 1) * 3 + 1).padStart(2, '0');
    return {
      isCalendar: true,
      calendarType: 'quarter',
      implicitDate: quarterlyMatch[1] + '-' + qMonth + '-01',
      implicitWeek: null,
      displayTitle: quarterlyMatch[1] + '-Q' + quarterlyMatch[2],
    };
  }

  // Yearly: YYYY
  var yearlyMatch = baseName.match(/^(\d{4})$/);
  if (yearlyMatch) {
    return {
      isCalendar: true,
      calendarType: 'year',
      implicitDate: yearlyMatch[1] + '-01-01',
      implicitWeek: null,
      displayTitle: yearlyMatch[1],
    };
  }

  return { isCalendar: false };
}

/**
 * Extract tasks from a single note and append to the tasks array.
 * isCalendarNote: if true, tasks inherit the note's date as their scheduled date.
 * calendarInfo: result of getCalendarNoteInfo() for calendar notes.
 */
function extractTasksFromNote(note, tasks, config, calendarInfo) {
  var filename = note.filename || '';

  // Check folder exclusions (for project notes)
  if (!calendarInfo || !calendarInfo.isCalendar) {
    var shouldExclude = false;
    for (var e = 0; e < config.foldersToExclude.length; e++) {
      var folder = config.foldersToExclude[e];
      if (filename.startsWith(folder + '/') || filename.startsWith(folder)) {
        shouldExclude = true;
        break;
      }
    }
    if (shouldExclude) return;
  }

  var isCalendar = calendarInfo && calendarInfo.isCalendar;
  var noteTitle;
  var noteFolder;

  if (isCalendar) {
    // Calendar notes: use date-based title and a virtual "Calendar" folder
    noteTitle = calendarInfo.displayTitle;
    noteFolder = '(calendar)';
  } else {
    noteTitle = note.title || filename.replace(/\.md$|\.txt$/, '');
    var folderParts = filename.split('/');
    noteFolder = folderParts.length > 1 ? folderParts.slice(0, -1).join('/') : '(root)';
  }

  var noteHashtags = note.hashtags || [];
  var noteMentions = note.mentions || [];

  var paras = note.paragraphs || [];
  var currentHeading = '';

  for (var i = 0; i < paras.length; i++) {
    var p = paras[i];

    if (p.type === 'title' && p.headingLevel && p.headingLevel >= 2) {
      currentHeading = p.content || '';
      continue;
    }

    // Determine if this is a task or checklist paragraph
    var taskKind = '';
    if (p.type === 'open' || p.type === 'done' || p.type === 'cancelled') {
      taskKind = 'task';
    } else if (p.type === 'checklist' || p.type === 'checklistDone' || p.type === 'checklistCancelled') {
      taskKind = 'checklist';
    } else {
      continue;
    }

    // Normalize type for uniform handling
    var normalizedType = p.type;
    if (p.type === 'checklist') normalizedType = 'open';
    else if (p.type === 'checklistDone') normalizedType = 'done';
    else if (p.type === 'checklistCancelled') normalizedType = 'cancelled';

    var content = p.content || '';

    // Parse priority
    var priority = 0;
    if (content.startsWith('!!! ')) priority = 3;
    else if (content.startsWith('!! ')) priority = 2;
    else if (content.startsWith('! ')) priority = 1;

    // Parse explicit scheduled date from task content
    var schedMatch = content.match(/>(\d{4}-\d{2}-\d{2})/);
    var weekMatch = content.match(/>(\d{4}-W\d{2})/);
    var scheduledDate = schedMatch ? schedMatch[1] : null;
    var scheduledWeek = weekMatch ? weekMatch[1] : null;

    // For calendar notes: if task has no explicit schedule, inherit from the note's date
    if (isCalendar && !scheduledDate && !scheduledWeek) {
      if (calendarInfo.implicitDate) {
        scheduledDate = calendarInfo.implicitDate;
      }
      if (calendarInfo.implicitWeek) {
        scheduledWeek = calendarInfo.implicitWeek;
      }
    }

    // Parse due date from @due(YYYY-MM-DD)
    var dueMatch = content.match(/@due\((\d{4}-\d{2}-\d{2})\)/);
    var dueDate = dueMatch ? dueMatch[1] : null;

    // Parse inline tags (#tag)
    var inlineTags = [];
    var tagRegex = /#[\w\-\/]+/g;
    var tm;
    while ((tm = tagRegex.exec(content)) !== null) {
      inlineTags.push(tm[0]);
    }

    // Parse inline mentions (@mention)
    var inlineMentions = [];
    var mentionRegex = /@[\w\-]+(?:\([^)]*\))?/g;
    var mm;
    while ((mm = mentionRegex.exec(content)) !== null) {
      // Skip @done and @due (system metadata), but keep @repeat (useful for filtering)
      if (!mm[0].startsWith('@done') && !mm[0].startsWith('@due')) {
        inlineMentions.push(mm[0].replace(/\([^)]*\)$/, ''));
      }
    }

    // Use only task-level inline tags for filtering (not note-level)
    var allTags = inlineTags.slice();

    // Clean display content
    var display = content;
    display = display.replace(/^!{1,3}\s*/, '');
    display = display.replace(/\s*>\d{4}-\d{2}-\d{2}(\s+\d{1,2}:\d{2}\s*(AM|PM)(\s*-\s*\d{1,2}:\d{2}\s*(AM|PM))?)?/gi, '');
    display = display.replace(/\s*>\d{4}-W\d{2}/g, '');
    display = display.replace(/\s*>today/g, '');
    display = display.replace(/\s*@done\([^)]*\)/g, '');
    display = display.replace(/\s*@repeat\([^)]*\)/g, '');

    tasks.push({
      lineIndex: i,
      filename: filename,
      encodedFilename: encSafe(filename),
      noteTitle: noteTitle,
      noteFolder: noteFolder,
      heading: currentHeading,
      content: display.trim(),
      rawContent: content,
      type: normalizedType,
      originalType: p.type,
      taskKind: taskKind,
      indentLevel: p.indentLevel || 0,
      priority: priority,
      scheduledDate: scheduledDate,
      scheduledWeek: scheduledWeek,
      dueDate: dueDate,
      tags: allTags,
      mentions: inlineMentions,
      noteTags: noteHashtags,
      isCalendarNote: isCalendar,
      calendarType: isCalendar ? calendarInfo.calendarType : null,
    });
  }
}

/**
 * Scan all notes (project + calendar) for tasks and return a flat array.
 * Calendar note tasks inherit their scheduled date from the note's date.
 */
function scanAllTasks() {
  var config = getSettings();
  var tasks = [];

  // Scan project notes
  var projectNotes = DataStore.projectNotes;
  for (var n = 0; n < projectNotes.length; n++) {
    extractTasksFromNote(projectNotes[n], tasks, config, null);
  }

  // Scan calendar notes
  var calendarNotes = DataStore.calendarNotes;
  for (var c = 0; c < calendarNotes.length; c++) {
    var calNote = calendarNotes[c];
    var calInfo = getCalendarNoteInfo(calNote);
    if (calInfo.isCalendar) {
      extractTasksFromNote(calNote, tasks, config, calInfo);
    }
  }

  return tasks;
}

// ============================================
// QUERY PARSER — Todoist-like syntax
// ============================================

/**
 * Parse a query string into a structured filter.
 * Syntax:
 *   - Free text: matches task content (case insensitive)
 *   - #tag: matches tasks with that tag
 *   - @mention: matches tasks with that mention
 *   - p1, p2, p3: priority filters (p1=highest/!!!, p2=medium/!!, p3=lowest/!)
 *   - open, done, cancelled: status filters
 *   - today, tomorrow, yesterday: scheduled date filters
 *   - this week, next week, this month, next month: date range filters
 *   - overdue: tasks with due/scheduled date before today
 *   - no date: tasks without any scheduled date
 *   - has date: tasks with a scheduled date
 *   - folder:Name: filter by folder (partial match)
 *   - note:Name: filter by note title (partial match)
 *   - checklist: filter to checklist items only (by default, only tasks are shown)
 *   - task, tasks: explicitly filter to tasks only (the default)
 *   - &, AND: AND combinator (default between tokens)
 *   - |, OR: OR combinator
 *   - NOT or ! prefix: negate next token (e.g. !#waiting, !@someone, !open)
 */
function parseQuery(queryStr) {
  if (!queryStr || !queryStr.trim()) return null;

  var tokens = tokenize(queryStr.trim());
  return buildFilterTree(tokens);
}

function tokenize(str) {
  var tokens = [];
  var i = 0;
  var len = str.length;

  while (i < len) {
    // Skip whitespace
    while (i < len && str[i] === ' ') i++;
    if (i >= len) break;

    var ch = str[i];

    // OR operator
    if (ch === '|') {
      tokens.push({ type: 'OR' });
      i++;
      continue;
    }

    // AND operator
    if (ch === '&') {
      tokens.push({ type: 'AND' });
      i++;
      continue;
    }

    // Quoted string
    if (ch === '"' || ch === "'") {
      var quote = ch;
      i++;
      var qstart = i;
      while (i < len && str[i] !== quote) i++;
      tokens.push({ type: 'TEXT', value: str.slice(qstart, i).toLowerCase() });
      if (i < len) i++; // skip closing quote
      continue;
    }

    // Handle ! prefix as NOT shorthand (e.g. !#tag, !@mention, !open)
    if (ch === '!' && i + 1 < len && str[i + 1] !== ' ' && str[i + 1] !== '!' ) {
      tokens.push({ type: 'NOT' });
      i++;
      continue;
    }

    // Read a word/token
    var wstart = i;
    while (i < len && str[i] !== ' ' && str[i] !== '|' && str[i] !== '&') i++;
    var word = str.slice(wstart, i);

    // Check for compound tokens
    var lw = word.toLowerCase();

    if (lw === 'and') { tokens.push({ type: 'AND' }); continue; }
    if (lw === 'or') { tokens.push({ type: 'OR' }); continue; }
    if (lw === 'not') { tokens.push({ type: 'NOT' }); continue; }

    // "this week", "next week", "this month", "next month", "no date", "has date"
    var rest = str.slice(wstart).toLowerCase();
    if (rest.startsWith('this week')) { tokens.push({ type: 'FILTER', filterType: 'dateRange', value: 'this week' }); i = wstart + 9; continue; }
    if (rest.startsWith('next week')) { tokens.push({ type: 'FILTER', filterType: 'dateRange', value: 'next week' }); i = wstart + 9; continue; }
    if (rest.startsWith('this month')) { tokens.push({ type: 'FILTER', filterType: 'dateRange', value: 'this month' }); i = wstart + 10; continue; }
    if (rest.startsWith('next month')) { tokens.push({ type: 'FILTER', filterType: 'dateRange', value: 'next month' }); i = wstart + 10; continue; }
    if (rest.startsWith('no date')) { tokens.push({ type: 'FILTER', filterType: 'noDate', value: true }); i = wstart + 7; continue; }
    if (rest.startsWith('has date')) { tokens.push({ type: 'FILTER', filterType: 'hasDate', value: true }); i = wstart + 8; continue; }

    // Priority shortcuts
    if (lw === 'p1') { tokens.push({ type: 'FILTER', filterType: 'priority', value: 3 }); continue; } // p1 = !!! (highest)
    if (lw === 'p2') { tokens.push({ type: 'FILTER', filterType: 'priority', value: 2 }); continue; } // p2 = !!
    if (lw === 'p3') { tokens.push({ type: 'FILTER', filterType: 'priority', value: 1 }); continue; } // p3 = ! (lowest)

    // Status keywords
    if (lw === 'open') { tokens.push({ type: 'FILTER', filterType: 'status', value: 'open' }); continue; }
    if (lw === 'done') { tokens.push({ type: 'FILTER', filterType: 'status', value: 'done' }); continue; }
    if (lw === 'cancelled' || lw === 'canceled') { tokens.push({ type: 'FILTER', filterType: 'status', value: 'cancelled' }); continue; }

    // Task kind keywords
    if (lw === 'checklist' || lw === 'checklists') { tokens.push({ type: 'FILTER', filterType: 'taskKind', value: 'checklist' }); continue; }
    if (lw === 'task' || lw === 'tasks') { tokens.push({ type: 'FILTER', filterType: 'taskKind', value: 'task' }); continue; }

    // Date keywords
    if (lw === 'today') { tokens.push({ type: 'FILTER', filterType: 'dateRange', value: 'today' }); continue; }
    if (lw === 'tomorrow') { tokens.push({ type: 'FILTER', filterType: 'dateRange', value: 'tomorrow' }); continue; }
    if (lw === 'yesterday') { tokens.push({ type: 'FILTER', filterType: 'dateRange', value: 'yesterday' }); continue; }
    if (lw === 'overdue') { tokens.push({ type: 'FILTER', filterType: 'overdue', value: true }); continue; }

    // Tag filter
    if (word.startsWith('#')) {
      tokens.push({ type: 'FILTER', filterType: 'tag', value: word });
      continue;
    }

    // Mention filter
    if (word.startsWith('@')) {
      tokens.push({ type: 'FILTER', filterType: 'mention', value: word });
      continue;
    }

    // folder: prefix
    if (lw.startsWith('folder:')) {
      tokens.push({ type: 'FILTER', filterType: 'folder', value: word.slice(7) });
      continue;
    }

    // note: prefix
    if (lw.startsWith('note:')) {
      tokens.push({ type: 'FILTER', filterType: 'note', value: word.slice(5) });
      continue;
    }

    // Free text
    tokens.push({ type: 'TEXT', value: lw });
  }

  return tokens;
}

function buildFilterTree(tokens) {
  if (!tokens || tokens.length === 0) return null;

  // Build a list of conditions connected by AND/OR
  var conditions = [];
  var currentOp = 'AND';
  var negate = false;

  for (var i = 0; i < tokens.length; i++) {
    var tok = tokens[i];

    if (tok.type === 'AND') { currentOp = 'AND'; continue; }
    if (tok.type === 'OR') { currentOp = 'OR'; continue; }
    if (tok.type === 'NOT') { negate = true; continue; }

    var condition;
    if (tok.type === 'FILTER') {
      condition = { type: 'filter', filterType: tok.filterType, value: tok.value };
    } else if (tok.type === 'TEXT') {
      condition = { type: 'text', value: tok.value };
    } else {
      continue;
    }

    if (negate) {
      condition = { type: 'not', child: condition };
      negate = false;
    }

    conditions.push({ op: currentOp, condition: condition });
    currentOp = 'AND'; // default back to AND
  }

  if (conditions.length === 0) return null;
  if (conditions.length === 1) return conditions[0].condition;

  // Build tree respecting OR as lower precedence than AND
  var orGroups = [];
  var currentAndGroup = [conditions[0].condition];

  for (var j = 1; j < conditions.length; j++) {
    if (conditions[j].op === 'OR') {
      orGroups.push(currentAndGroup);
      currentAndGroup = [conditions[j].condition];
    } else {
      currentAndGroup.push(conditions[j].condition);
    }
  }
  orGroups.push(currentAndGroup);

  // Build AND nodes for each group
  var orNodes = [];
  for (var g = 0; g < orGroups.length; g++) {
    var group = orGroups[g];
    if (group.length === 1) {
      orNodes.push(group[0]);
    } else {
      orNodes.push({ type: 'and', children: group });
    }
  }

  // Combine OR nodes
  if (orNodes.length === 1) return orNodes[0];
  return { type: 'or', children: orNodes };
}

// ============================================
// FILTER EVALUATION
// ============================================

function evaluateFilter(task, filter) {
  if (!filter) return true;

  switch (filter.type) {
    case 'text':
      return task.content.toLowerCase().indexOf(filter.value) !== -1 ||
             task.rawContent.toLowerCase().indexOf(filter.value) !== -1 ||
             task.noteTitle.toLowerCase().indexOf(filter.value) !== -1;

    case 'not':
      return !evaluateFilter(task, filter.child);

    case 'and':
      for (var a = 0; a < filter.children.length; a++) {
        if (!evaluateFilter(task, filter.children[a])) return false;
      }
      return true;

    case 'or':
      for (var o = 0; o < filter.children.length; o++) {
        if (evaluateFilter(task, filter.children[o])) return true;
      }
      return false;

    case 'filter':
      return evaluateFilterCondition(task, filter);

    default:
      return true;
  }
}

function evaluateFilterCondition(task, filter) {
  var today = getTodayStr();

  switch (filter.filterType) {
    case 'priority':
      return task.priority === filter.value;

    case 'status':
      return task.type === filter.value;

    case 'taskKind':
      return task.taskKind === filter.value;

    case 'tag': {
      var tagVal = filter.value.toLowerCase();
      // Bare "#" means "has any tag"
      if (tagVal === '#') return task.tags.length > 0;
      for (var t = 0; t < task.tags.length; t++) {
        if (task.tags[t].toLowerCase() === tagVal ||
            task.tags[t].toLowerCase().startsWith(tagVal)) return true;
      }
      return false;
    }

    case 'mention': {
      var mentionVal = filter.value.toLowerCase();
      // Bare "@" means "has any mention"
      if (mentionVal === '@') return task.mentions.length > 0;
      for (var m = 0; m < task.mentions.length; m++) {
        if (task.mentions[m].toLowerCase() === mentionVal ||
            task.mentions[m].toLowerCase().startsWith(mentionVal)) return true;
      }
      return false;
    }

    case 'folder':
      return task.noteFolder.toLowerCase().indexOf(filter.value.toLowerCase()) !== -1;

    case 'note':
      return task.noteTitle.toLowerCase().indexOf(filter.value.toLowerCase()) !== -1;

    case 'noDate':
      return !task.scheduledDate && !task.scheduledWeek && !task.dueDate;

    case 'hasDate':
      return !!(task.scheduledDate || task.scheduledWeek || task.dueDate);

    case 'overdue': {
      var effectiveDate = task.dueDate || task.scheduledDate;
      if (!effectiveDate) return false;
      return effectiveDate < today;
    }

    case 'dateRange':
      return matchDateRange(task, filter.value);

    default:
      return true;
  }
}

function matchDateRange(task, rangeStr) {
  var effectiveDate = task.scheduledDate || task.dueDate;

  switch (rangeStr) {
    case 'today':
      return effectiveDate === getTodayStr() || (task.scheduledWeek && task.scheduledWeek === getISOWeek(getTodayStr()));
    case 'tomorrow':
      return effectiveDate === getTomorrowStr();
    case 'yesterday':
      return effectiveDate === getYesterdayStr();
    case 'this week': {
      if (!effectiveDate) return false;
      return effectiveDate >= getStartOfWeek() && effectiveDate <= getEndOfWeek();
    }
    case 'next week': {
      if (!effectiveDate) return false;
      return effectiveDate >= getStartOfNextWeek() && effectiveDate <= getEndOfNextWeek();
    }
    case 'this month': {
      if (!effectiveDate) return false;
      return effectiveDate >= getStartOfMonth() && effectiveDate <= getEndOfMonth();
    }
    case 'next month': {
      if (!effectiveDate) return false;
      return effectiveDate >= getStartOfNextMonth() && effectiveDate <= getEndOfNextMonth();
    }
    default:
      return false;
  }
}

// ============================================
// GROUPING
// ============================================

function groupTasks(tasks, groupBy) {
  var groups = {};
  var groupOrder = [];

  for (var i = 0; i < tasks.length; i++) {
    var task = tasks[i];
    var keys = getGroupKeys(task, groupBy);

    for (var k = 0; k < keys.length; k++) {
      var key = keys[k];
      if (!groups[key]) {
        groups[key] = [];
        groupOrder.push(key);
      }
      groups[key].push(task);
    }
  }

  // Sort groups
  groupOrder.sort(function(a, b) {
    if (groupBy === 'priority') {
      var priOrder = { 'Highest Priority': 3, 'High Priority': 2, 'Priority': 1, 'No priority': 0 };
      return (priOrder[b] || 0) - (priOrder[a] || 0);
    }
    if (groupBy === 'status') {
      var statusOrder = { 'Open': 0, 'Done': 1, 'Cancelled': 2 };
      return (statusOrder[a] || 0) - (statusOrder[b] || 0);
    }
    if (groupBy === 'date') {
      if (a === 'No date') return 1;
      if (b === 'No date') return -1;
      if (a === 'Overdue') return -1;
      if (b === 'Overdue') return 1;
      if (a === 'Today') return -1;
      if (b === 'Today') return 1;
      if (a === 'Tomorrow') return (b === 'Today' || b === 'Overdue') ? 1 : -1;
      if (b === 'Tomorrow') return (a === 'Today' || a === 'Overdue') ? -1 : 1;
      return a.localeCompare(b);
    }
    return a.localeCompare(b);
  });

  var result = [];
  for (var g = 0; g < groupOrder.length; g++) {
    result.push({ label: groupOrder[g], tasks: groups[groupOrder[g]] });
  }
  return result;
}

function getGroupKeys(task, groupBy) {
  switch (groupBy) {
    case 'folder':
      return [task.noteFolder || '(root)'];
    case 'note':
      return [task.noteTitle];
    case 'status':
      return [task.type === 'open' ? 'Open' : task.type === 'done' ? 'Done' : 'Cancelled'];
    case 'tag':
      if (task.tags.length === 0) return ['No tags'];
      return task.tags;
    case 'mention':
      if (task.mentions.length === 0) return ['No mentions'];
      return task.mentions;
    case 'date': {
      var d = task.scheduledDate || task.dueDate;
      if (!d) return ['No date'];
      if (d < getTodayStr()) return ['Overdue'];
      if (d === getTodayStr()) return ['Today'];
      if (d === getTomorrowStr()) return ['Tomorrow'];
      return [d];
    }
    case 'priority':
      if (task.priority === 0) return ['No priority'];
      if (task.priority === 3) return ['Highest Priority'];
      if (task.priority === 2) return ['High Priority'];
      return ['Priority'];
    case 'none':
      return ['All tasks'];
    default:
      return [task.noteTitle];
  }
}

// ============================================
// HTML GENERATION
// ============================================

function esc(str) {
  if (!str) return '';
  return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/**
 * Render basic markdown to HTML.
 * Supports: **bold**, *italic*, `inline code`, ~~strikethrough~~, [links](url)
 * First escapes HTML, then applies markdown transforms.
 */
function renderMarkdown(str) {
  if (!str) return '';
  // Escape HTML first
  var s = str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

  // Wiki links: [[Note Name]] → clickable link to open in NotePlan split view
  s = s.replace(/\[\[([^\]]+)\]\]/g, function(match, noteName) {
    var encoded = encodeURIComponent(noteName);
    var url = 'noteplan://x-callback-url/openNote?noteTitle=' + encoded + '&amp;splitView=yes';
    return '<a class="tz-md-link tz-wiki-link" href="' + url + '" title="' + noteName.replace(/"/g, '&amp;quot;') + '">' + noteName + '</a>';
  });

  // Links: [text](url) — must come before bold/italic to avoid conflicts
  s = s.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a class="tz-md-link" href="$2" title="$2">$1</a>');

  // Inline code: `code`
  s = s.replace(/`([^`]+)`/g, '<code class="tz-md-code">$1</code>');

  // Bold + italic: ***text*** or ___text___
  s = s.replace(/\*\*\*([^*]+)\*\*\*/g, '<strong><em>$1</em></strong>');

  // Bold: **text** or __text__
  s = s.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
  s = s.replace(/__([^_]+)__/g, '<strong>$1</strong>');

  // Italic: *text* or _text_ (but not mid-word underscores)
  s = s.replace(/\*([^*]+)\*/g, '<em>$1</em>');
  s = s.replace(/(?<!\w)_([^_]+)_(?!\w)/g, '<em>$1</em>');

  // Strikethrough: ~~text~~
  s = s.replace(/~~([^~]+)~~/g, '<del>$1</del>');

  // ==highlight== (NotePlan supports this)
  s = s.replace(/==([^=]+)==/g, '<mark class="tz-md-highlight">$1</mark>');

  // Tags: #tag (but not inside HTML attributes/tags already rendered)
  s = s.replace(/(^|[\s(])#([\w][\w/-]*)/g, '$1<span class="tz-tag">#$2</span>');

  // Mentions: @mention
  s = s.replace(/(^|[\s(])@([\w][\w/-]*(?:\([^)]*\))?)/g, '$1<span class="tz-mention">@$2</span>');

  return s;
}

/**
 * Convert NotePlan's #AARRGGBB hex to standard #RRGGBBAA (or pass through #RRGGBB)
 */
function npColor(c) {
  if (!c) return null;
  if (c.match && c.match(/^#[0-9A-Fa-f]{8}$/)) {
    return '#' + c.slice(3, 9) + c.slice(1, 3);
  }
  return c;
}

function isLightTheme() {
  try {
    var theme = Editor.currentTheme;
    if (!theme) return false;
    if (theme.mode === 'light') return true;
    if (theme.mode === 'dark') return false;
    // Fallback: check luminance of background
    var vals = theme.values || {};
    var bg = npColor((vals.editor || {}).backgroundColor);
    if (bg) {
      var m = bg.match(/^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})/i);
      if (m) {
        var lum = (parseInt(m[1], 16) * 299 + parseInt(m[2], 16) * 587 + parseInt(m[3], 16) * 114) / 1000;
        return lum > 140;
      }
    }
  } catch (e) {}
  return false;
}

function getThemeCSS() {
  try {
    var theme = Editor.currentTheme;
    if (!theme) return '';
    var vals = theme.values || {};
    var editor = vals.editor || {};
    var styles = [];
    var bg = npColor(editor.backgroundColor);
    var altBg = npColor(editor.altBackgroundColor);
    var text = npColor(editor.textColor);
    var tint = npColor(editor.tintColor);
    if (bg) styles.push('--bg-main-color: ' + bg);
    if (altBg) styles.push('--bg-alt-color: ' + altBg);
    if (text) styles.push('--fg-main-color: ' + text);
    if (tint) styles.push('--tint-color: ' + tint);
    if (styles.length > 0) return ':root { ' + styles.join('; ') + '; }';
  } catch (e) {}
  return '';
}

function buildFilterSidebar(savedFilters, activeFilterId) {
  var html = '<div class="tz-sidebar">';
  html += '<div class="tz-sidebar-header">';
  html += '<span class="tz-sidebar-title">Filters</span>';
  html += '</div>';

  // Built-in quick filters
  html += '<div class="tz-filter-group">';
  html += '<div class="tz-filter-group-label">Quick Filters</div>';
  var builtins = [
    { id: '__overdue', icon: 'fa-solid fa-clock', label: 'Overdue', query: 'open overdue' },
    { id: '__high', icon: 'fa-solid fa-exclamation', label: 'Priority', query: 'open p1 | open p2 | open p3' },
    { id: '__today', icon: 'fa-regular fa-calendar', label: 'Today', query: 'open today' },
    { id: '__thisweek', icon: 'fa-solid fa-calendar-week', label: 'This Week', query: 'open this week' },
    { id: '__nodate', icon: 'fa-solid fa-calendar-xmark', label: 'No Date', query: 'open no date' },
    { id: '__all', icon: 'fa-solid fa-list-check', label: 'All Open', query: 'open' },
  ];

  for (var b = 0; b < builtins.length; b++) {
    var bi = builtins[b];
    var active = activeFilterId === bi.id ? ' active' : '';
    html += '<button class="tz-filter-item' + active + '" data-filter-id="' + bi.id + '" data-query="' + esc(bi.query) + '">';
    html += '<i class="' + bi.icon + '"></i>';
    html += '<span>' + esc(bi.label) + '</span>';
    html += '</button>';
  }
  html += '</div>';

  // Saved filters
  if (savedFilters.length > 0) {
    html += '<div class="tz-filter-group">';
    html += '<div class="tz-filter-group-label">Saved Filters</div>';
    for (var s = 0; s < savedFilters.length; s++) {
      var sf = savedFilters[s];
      var active2 = activeFilterId === sf.id ? ' active' : '';
      html += '<button class="tz-filter-item saved' + active2 + '" data-filter-id="' + esc(sf.id) + '" data-query="' + esc(sf.query) + '" data-filter-name="' + esc(sf.name) + '" draggable="true">';
      html += '<span class="tz-filter-name">' + esc(sf.name) + '</span>';
      html += '<i class="fa-solid fa-grip-vertical tz-drag-handle"></i>';
      html += '</button>';
    }
    html += '<button class="tz-filter-item tz-new-filter" data-action="newFilter">';
    html += '<i class="fa-solid fa-plus" style="font-size:9px;width:14px;text-align:center;opacity:0.5"></i>';
    html += '<span>New filter...</span>';
    html += '</button>';
    html += '</div>';
  }

  html += '</div>';
  return html;
}

function buildMainContent(tasks, groupBy, activeQuery, filterOriginalQuery, totalScanned) {
  var html = '<div class="tz-main">';

  // Header with search bar
  html += '<div class="tz-header">';
  html += '<button class="tz-sidebar-toggle" data-action="toggleSidebar"><i class="fa-solid fa-bars"></i></button>';
  html += '<div class="tz-search-wrap">';
  html += '<i class="fa-solid fa-magnifying-glass tz-search-icon"></i>';
  html += '<input class="tz-search-input" type="text" placeholder="Filter: #tag, @mention, p1-p3, today, folder:Name, open|done..." value="' + esc(activeQuery) + '" data-original-query="' + esc(filterOriginalQuery) + '">';
  html += '<button class="tz-search-clear" data-tooltip="Clear"><i class="fa-solid fa-xmark"></i></button>';
  html += '</div>';
  html += '<div class="tz-save-area" style="display:none;">';
  html += '<button class="tz-btn tz-save-btn"><i class="fa-solid fa-bookmark"></i> Save</button>';
  html += '</div>';
  html += '</div>';

  // Group-by toolbar
  html += '<div class="tz-toolbar">';
  html += '<span class="tz-toolbar-label">Group by</span>';
  var groupOptions = ['note', 'folder', 'status', 'tag', 'mention', 'date', 'priority', 'none'];
  for (var g = 0; g < groupOptions.length; g++) {
    var gopt = groupOptions[g];
    var gactive = groupBy === gopt ? ' active' : '';
    html += '<button class="tz-group-btn' + gactive + '" data-group="' + gopt + '">' + gopt.charAt(0).toUpperCase() + gopt.slice(1) + '</button>';
  }
  html += '<span class="tz-toolbar-count">' + tasks.length + ' tasks' + (totalScanned > 0 ? ' / ' + totalScanned + ' scanned' : '') + '</span>';
  html += '</div>';

  // Task list
  html += '<div class="tz-body">';

  if (tasks.length === 0) {
    html += '<div class="tz-empty">';
    html += '<div class="tz-empty-icon"><i class="fa-solid fa-filter-circle-xmark"></i></div>';
    html += '<div class="tz-empty-title">No tasks match this filter</div>';
    html += '<div class="tz-empty-desc">Try adjusting your search query or filter settings.</div>';
    html += '</div>';
  } else {
    var groups = groupTasks(tasks, groupBy);
    for (var gi = 0; gi < groups.length; gi++) {
      var grp = groups[gi];
      html += '<div class="tz-group">';
      html += '<div class="tz-group-header">';
      html += '<span class="tz-group-label">' + esc(grp.label) + '</span>';
      html += '<span class="tz-group-count">' + grp.tasks.length + '</span>';
      html += '</div>';
      html += '<div class="tz-task-list">';
      for (var ti = 0; ti < grp.tasks.length; ti++) {
        html += buildTaskRow(grp.tasks[ti], groupBy);
      }
      html += '</div>';
      html += '</div>';
    }
  }

  html += '</div>'; // .tz-body
  html += '</div>'; // .tz-main
  return html;
}

function buildTaskRow(task, groupBy) {
  var isChecklist = task.taskKind === 'checklist';
  var cbClass = 'tz-task-cb ' + task.type + (isChecklist ? ' checklist' : '');
  var cbIcon;
  if (isChecklist) {
    if (task.type === 'done') cbIcon = 'fa-solid fa-square-check';
    else if (task.type === 'cancelled') cbIcon = 'fa-solid fa-square-minus';
    else cbIcon = 'fa-regular fa-square';
  } else {
    if (task.type === 'done') cbIcon = 'fa-solid fa-circle-check';
    else if (task.type === 'cancelled') cbIcon = 'fa-solid fa-circle-minus';
    else cbIcon = 'fa-regular fa-circle';
  }

  var rowClass = 'tz-task';
  if (task.type === 'done') rowClass += ' is-done';
  if (task.type === 'cancelled') rowClass += ' is-cancelled';
  if (task.indentLevel > 0) rowClass += ' indent-' + Math.min(task.indentLevel, 3);

  var html = '<div class="' + rowClass + '" data-line-index="' + task.lineIndex + '" data-encoded-filename="' + task.encodedFilename + '">';

  // Checkbox
  html += '<span class="' + cbClass + '" data-action="toggleComplete"><i class="' + cbIcon + '"></i></span>';

  // Priority badge
  if (task.priority > 0) {
    html += '<span class="tz-task-pri p' + task.priority + '" data-action="cyclePriority">' + Array(task.priority + 1).join('!') + '</span>';
  }

  // Content
  html += '<span class="tz-task-content">' + renderMarkdown(task.content) + '</span>';

  // Meta badges
  html += '<span class="tz-task-meta">';

  // Note badge (hide if grouped by note)
  if (groupBy !== 'note') {
    var noteIcon = task.isCalendarNote ? 'fa-regular fa-calendar-days' : 'fa-regular fa-note-sticky';
    html += '<span class="tz-task-badge note" data-action="openNote" data-tooltip="' + esc(task.noteTitle) + '"><i class="' + noteIcon + '"></i> ' + esc(truncate(task.noteTitle, 20)) + '</span>';
  }

  // Schedule badge with date-aware coloring
  if (task.scheduledDate) {
    var dateClass = 'date';
    var todayStr = getTodayStr();
    var weekStart = getStartOfWeek();
    var weekEnd = getEndOfWeek();
    if (task.type === 'open') {
      if (task.scheduledDate < todayStr) dateClass += ' overdue';
      else if (task.scheduledDate === todayStr) dateClass += ' today';
      else if (task.scheduledDate >= weekStart && task.scheduledDate <= weekEnd) dateClass += ' this-week';
    }
    html += '<span class="tz-task-badge ' + dateClass + '" data-action="showSchedule"><i class="fa-regular fa-calendar"></i> ' + esc(formatDate(task.scheduledDate)) + '</span>';
  } else if (task.scheduledWeek) {
    var weekClass = 'date';
    if (task.type === 'open') {
      var currentWeek = getISOWeek(getTodayStr());
      if (task.scheduledWeek < currentWeek) weekClass += ' overdue';
      else if (task.scheduledWeek === currentWeek) weekClass += ' this-week';
    }
    html += '<span class="tz-task-badge ' + weekClass + '" data-action="showSchedule"><i class="fa-regular fa-calendar"></i> ' + esc(task.scheduledWeek) + '</span>';
  }

  // Due date badge
  if (task.dueDate) {
    var dueClass = task.dueDate < getTodayStr() ? ' overdue' : '';
    html += '<span class="tz-task-badge due' + dueClass + '"><i class="fa-solid fa-flag"></i> ' + esc(formatDate(task.dueDate)) + '</span>';
  }

  // Folder badge (hide if grouped by folder)
  if (groupBy !== 'folder' && task.noteFolder !== '(root)') {
    html += '<span class="tz-task-badge folder"><i class="fa-solid fa-folder"></i> ' + esc(truncate(task.noteFolder, 15)) + '</span>';
  }

  html += '</span>'; // .tz-task-meta

  // Hover actions
  html += '<span class="tz-task-acts">';
  if (task.priority === 0) {
    html += '<button class="tz-task-act" data-action="cyclePriority" data-tooltip="Set priority"><i class="fa-solid fa-exclamation"></i></button>';
  }
  if (!task.scheduledDate && !task.scheduledWeek) {
    html += '<button class="tz-task-act" data-action="showSchedule" data-tooltip="Schedule"><i class="fa-regular fa-calendar"></i></button>';
  }
  html += '<button class="tz-task-act" data-action="assignPerson" data-tooltip="Assign"><i class="fa-solid fa-user-plus"></i></button>';
  html += '<button class="tz-task-act" data-action="openNote" data-tooltip="Open note"><i class="fa-solid fa-arrow-up-right-from-square"></i></button>';
  if (task.type !== 'cancelled') {
    html += '<button class="tz-task-act cancel" data-action="toggleCancel" data-tooltip="Cancel"><i class="fa-solid fa-xmark"></i></button>';
  }
  html += '</span>';

  html += '</div>';
  return html;
}

function truncate(str, max) {
  if (!str) return '';
  if (str.length <= max) return str;
  return str.slice(0, max) + '\u2026';
}

// Task cache — avoid re-scanning on every filter switch
var _taskCache = null;
// Filter result cache — cache filtered tasks by query (grouping is cheap)
var _filterResultCache = {};

function getCachedTasks() {
  if (_taskCache) return _taskCache;
  _taskCache = scanAllTasks();
  return _taskCache;
}

function invalidateTaskCache() {
  _taskCache = null;
  _filterResultCache = {};
}

function getFilterCacheKey(query) {
  return query || '';
}

function buildDashboardHTML(config, activeQuery, activeFilterId, groupBy) {
  var allTasks = getCachedTasks();
  var totalScanned = allTasks.length;

  // Collect all unique mentions for the assign picker
  var allMentionsSet = {};
  for (var mi = 0; mi < allTasks.length; mi++) {
    for (var mj = 0; mj < allTasks[mi].mentions.length; mj++) {
      var m = allTasks[mi].mentions[mj];
      if (m) allMentionsSet[m.toLowerCase()] = m;
    }
  }
  var allMentionsList = Object.keys(allMentionsSet).sort().map(function(k) { return allMentionsSet[k]; });

  // Check if query explicitly mentions task kind
  var queryLower = (activeQuery || '').toLowerCase();
  var hasKindKeyword = /\b(checklist|checklists|task|tasks)\b/.test(queryLower);

  var filter = parseQuery(activeQuery);
  var filteredTasks = [];
  for (var i = 0; i < allTasks.length; i++) {
    // If query doesn't specify task kind, default to tasks only
    if (!hasKindKeyword && allTasks[i].taskKind !== 'task') continue;
    if (evaluateFilter(allTasks[i], filter)) {
      filteredTasks.push(allTasks[i]);
    }
  }

  // Determine the filter's stored query (to detect user edits in the UI)
  var builtinQueries = {
    '__overdue': 'open overdue', '__high': 'open p1 | open p2 | open p3',
    '__today': 'open today', '__thisweek': 'open this week',
    '__nodate': 'open no date', '__all': 'open',
  };
  var filterOriginalQuery = activeQuery;
  if (activeFilterId && builtinQueries[activeFilterId]) {
    filterOriginalQuery = builtinQueries[activeFilterId];
  } else if (activeFilterId) {
    for (var fj = 0; fj < config.savedFilters.length; fj++) {
      if (config.savedFilters[fj].id === activeFilterId) { filterOriginalQuery = config.savedFilters[fj].query; break; }
    }
  }

  var html = '<script>var allMentions = ' + JSON.stringify(allMentionsList) + ';<\/script>\n';
  html += '<div class="tz-layout">';
  html += buildFilterSidebar(config.savedFilters, activeFilterId);
  html += '<div class="tz-sidebar-backdrop"></div>';
  html += buildMainContent(filteredTasks, groupBy, activeQuery, filterOriginalQuery, totalScanned);
  html += '</div>';
  return html;
}

function buildFullHTML(bodyContent) {
  var themeCSS = getThemeCSS();
  var pluginCSS = getInlineCSS();

  var faLinks = '\n' +
    '    <link href="../np.Shared/fontawesome.css" rel="stylesheet">\n' +
    '    <link href="../np.Shared/regular.min.flat4NP.css" rel="stylesheet">\n' +
    '    <link href="../np.Shared/solid.min.flat4NP.css" rel="stylesheet">\n';

  var themeAttr = isLightTheme() ? 'light' : 'dark';
  return '<!DOCTYPE html>\n<html data-theme="' + themeAttr + '">\n<head>\n' +
    '  <meta charset="utf-8">\n' +
    '  <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no, maximum-scale=1, viewport-fit=cover">\n' +
    '  <title>Task Zoom</title>\n' +
    faLinks +
    '  <style>' + themeCSS + '\n' + pluginCSS + '</style>\n' +
    '</head>\n<body>\n' +
    bodyContent + '\n' +
    '  <script>\n    var receivingPluginID = \'' + PLUGIN_ID + '\';\n  <\/script>\n' +
    '  <script type="text/javascript" src="../np.Shared/pluginToHTMLCommsBridge.js"><\/script>\n' +
    '  <script type="text/javascript" src="taskZoomEvents.js"><\/script>\n' +
    '</body>\n</html>';
}

// ============================================
// INLINE CSS
// ============================================

/**
 * Convert NotePlan's #AARRGGBB or #RRGGBB hex color to CSS rgba() string.
 */
function npColorToCSS(hex) {
  if (!hex || typeof hex !== 'string') return null;
  hex = hex.replace(/^#/, '');
  if (hex.length === 8) {
    // #AARRGGBB
    var a = parseInt(hex.substring(0, 2), 16) / 255;
    var r = parseInt(hex.substring(2, 4), 16);
    var g = parseInt(hex.substring(4, 6), 16);
    var b = parseInt(hex.substring(6, 8), 16);
    return 'rgba(' + r + ',' + g + ',' + b + ',' + a.toFixed(2) + ')';
  }
  if (hex.length === 6) {
    return '#' + hex;
  }
  return null;
}

/**
 * Read priority (flagged) colors from the current NotePlan theme.
 * Returns { pri3: { bg, color }, pri2: { bg, color }, pri1: { bg, color } }
 * where pri3 = !!! (highest), pri2 = !!, pri1 = ! (lowest).
 * Falls back to hardcoded defaults if theme data is unavailable.
 */
function getThemePriorityColors() {
  var defaults = {
    pri3: { bg: 'rgba(255,85,85,0.67)', color: '#FFB5B5' },
    pri2: { bg: 'rgba(255,85,85,0.47)', color: '#FFCCCC' },
    pri1: { bg: 'rgba(255,85,85,0.27)', color: '#FFDBBE' },
  };
  try {
    if (typeof Editor === 'undefined' || !Editor.currentTheme || !Editor.currentTheme.values) return defaults;
    var styles = Editor.currentTheme.values.styles || {};
    var f1 = styles['flagged-1']; // ! (one mark, lowest)
    var f2 = styles['flagged-2']; // !!
    var f3 = styles['flagged-3']; // !!! (three marks, highest)
    return {
      pri1: {
        bg: (f1 && f1.backgroundColor) ? npColorToCSS(f1.backgroundColor) || defaults.pri1.bg : defaults.pri1.bg,
        color: (f1 && f1.color) ? npColorToCSS(f1.color) || defaults.pri1.color : defaults.pri1.color,
      },
      pri2: {
        bg: (f2 && f2.backgroundColor) ? npColorToCSS(f2.backgroundColor) || defaults.pri2.bg : defaults.pri2.bg,
        color: (f2 && f2.color) ? npColorToCSS(f2.color) || defaults.pri2.color : defaults.pri2.color,
      },
      pri3: {
        bg: (f3 && f3.backgroundColor) ? npColorToCSS(f3.backgroundColor) || defaults.pri3.bg : defaults.pri3.bg,
        color: (f3 && f3.color) ? npColorToCSS(f3.color) || defaults.pri3.color : defaults.pri3.color,
      },
    };
  } catch (e) {
    return defaults;
  }
}

function priCSS(className) {
  var c = getThemePriorityColors();
  return '.' + className + '.p3 { background: ' + c.pri3.bg + '; color: ' + c.pri3.color + '; }\n' +
         '.' + className + '.p2 { background: ' + c.pri2.bg + '; color: ' + c.pri2.color + '; }\n' +
         '.' + className + '.p1 { background: ' + c.pri1.bg + '; color: ' + c.pri1.color + '; }\n';
}

function getInlineCSS() {
  return '\n' +
'/* ---- Dark theme (default) ---- */\n' +
':root, [data-theme="dark"] {\n' +
'  --tz-bg: var(--bg-main-color, #1a1a2e);\n' +
'  --tz-bg-card: var(--bg-alt-color, #16213e);\n' +
'  --tz-bg-elevated: color-mix(in srgb, var(--tz-bg-card) 85%, white 15%);\n' +
'  --tz-text: var(--fg-main-color, #e0e0e0);\n' +
'  --tz-text-muted: color-mix(in srgb, var(--tz-text) 55%, transparent);\n' +
'  --tz-text-faint: color-mix(in srgb, var(--tz-text) 35%, transparent);\n' +
'  --tz-accent: var(--tint-color, #F59E0B);\n' +
'  --tz-accent-soft: color-mix(in srgb, var(--tz-accent) 15%, transparent);\n' +
'  --tz-border: color-mix(in srgb, var(--tz-text) 10%, transparent);\n' +
'  --tz-border-strong: color-mix(in srgb, var(--tz-text) 18%, transparent);\n' +
'  --tz-green: #10B981;\n' +
'  --tz-green-soft: color-mix(in srgb, #10B981 12%, transparent);\n' +
'  --tz-yellow: #F59E0B;\n' +
'  --tz-yellow-soft: color-mix(in srgb, #F59E0B 12%, transparent);\n' +
'  --tz-red: #EF4444;\n' +
'  --tz-red-soft: color-mix(in srgb, #EF4444 12%, transparent);\n' +
'  --tz-blue: #3B82F6;\n' +
'  --tz-blue-soft: color-mix(in srgb, #3B82F6 12%, transparent);\n' +
'  --tz-purple: #8B5CF6;\n' +
'  --tz-purple-soft: color-mix(in srgb, #8B5CF6 12%, transparent);\n' +
'  --tz-orange: #F97316;\n' +
'  --tz-orange-soft: color-mix(in srgb, #F97316 12%, transparent);\n' +
'  --tz-gray: color-mix(in srgb, var(--tz-text) 40%, transparent);\n' +
'  --tz-radius: 10px;\n' +
'  --tz-radius-sm: 6px;\n' +
'  --tz-radius-xs: 4px;\n' +
'  --tz-sidebar-width: 200px;\n' +
'}\n' +
'/* ---- Light theme overrides ---- */\n' +
'[data-theme="light"] {\n' +
'  --tz-bg-elevated: color-mix(in srgb, var(--tz-bg-card) 92%, black 8%);\n' +
'  --tz-text-muted: color-mix(in srgb, var(--tz-text) 60%, transparent);\n' +
'  --tz-text-faint: color-mix(in srgb, var(--tz-text) 40%, transparent);\n' +
'  --tz-border: color-mix(in srgb, var(--tz-text) 12%, transparent);\n' +
'  --tz-border-strong: color-mix(in srgb, var(--tz-text) 22%, transparent);\n' +
'  --tz-green-soft: color-mix(in srgb, #10B981 10%, white);\n' +
'  --tz-yellow-soft: color-mix(in srgb, #F59E0B 10%, white);\n' +
'  --tz-red-soft: color-mix(in srgb, #EF4444 10%, white);\n' +
'  --tz-blue-soft: color-mix(in srgb, #3B82F6 10%, white);\n' +
'  --tz-purple-soft: color-mix(in srgb, #8B5CF6 10%, white);\n' +
'  --tz-orange-soft: color-mix(in srgb, #F97316 10%, white);\n' +
'  --tz-accent-soft: color-mix(in srgb, var(--tz-accent) 12%, white);\n' +
'}\n' +
'* { box-sizing: border-box; margin: 0; padding: 0; }\n' +
'body {\n' +
'  font-family: -apple-system, BlinkMacSystemFont, "SF Pro Text", "Segoe UI", system-ui, sans-serif;\n' +
'  background: var(--tz-bg); color: var(--tz-text);\n' +
'  font-size: 13px; line-height: 1.5;\n' +
'  -webkit-font-smoothing: antialiased; overflow: hidden; height: 100vh;\n' +
'}\n' +
'\n/* ---- Layout ---- */\n' +
'.tz-layout { display: flex; height: 100vh; overflow: hidden; }\n' +
'\n/* ---- Sidebar ---- */\n' +
'.tz-sidebar {\n' +
'  width: var(--tz-sidebar-width); flex-shrink: 0;\n' +
'  background: var(--tz-bg-card); border-right: 1px solid var(--tz-border);\n' +
'  display: flex; flex-direction: column; overflow-y: auto;\n' +
'}\n' +
'.tz-sidebar-header {\n' +
'  display: flex; align-items: center; justify-content: space-between;\n' +
'  padding: 12px 12px 8px; border-bottom: 1px solid var(--tz-border);\n' +
'}\n' +
'.tz-sidebar-title { font-size: 13px; font-weight: 700; letter-spacing: -0.01em; }\n' +
'.tz-sidebar-btn {\n' +
'  width: 24px; height: 24px; display: flex; align-items: center; justify-content: center;\n' +
'  border-radius: var(--tz-radius-xs); border: none; background: transparent;\n' +
'  color: var(--tz-text-muted); cursor: pointer; font-size: 11px;\n' +
'}\n' +
'.tz-sidebar-btn:hover { background: var(--tz-border); color: var(--tz-text); }\n' +
'.tz-filter-group { padding: 8px 8px 4px; }\n' +
'.tz-filter-group-label {\n' +
'  font-size: 10px; font-weight: 700; text-transform: uppercase;\n' +
'  letter-spacing: 0.06em; color: var(--tz-text-faint);\n' +
'  padding: 0 4px 4px; margin-bottom: 2px;\n' +
'}\n' +
'.tz-filter-item {\n' +
'  display: flex; align-items: center; gap: 8px; width: 100%;\n' +
'  padding: 5px 8px; font-size: 12px; font-weight: 500;\n' +
'  border-radius: var(--tz-radius-sm); border: none; background: transparent;\n' +
'  color: var(--tz-text-muted); cursor: pointer; text-align: left;\n' +
'  transition: all 0.12s ease; position: relative;\n' +
'}\n' +
'.tz-filter-item:hover { background: var(--tz-border); color: var(--tz-text); }\n' +
'.tz-filter-item.active { background: var(--tz-accent-soft); color: var(--tz-accent); font-weight: 600; }\n' +
'.tz-filter-item i { font-size: 11px; width: 14px; text-align: center; flex-shrink: 0; }\n' +
'.tz-filter-item span { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }\n' +
'.tz-filter-name { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; flex: 1; }\n' +
'\n/* ---- Drag Handle & Drag States ---- */\n' +
'.tz-drag-handle {\n' +
'  font-size: 9px; flex-shrink: 0; margin-left: auto;\n' +
'  color: var(--tz-text-faint); cursor: grab; opacity: 0;\n' +
'  transition: opacity 0.12s ease;\n' +
'}\n' +
'.tz-filter-item.saved:hover .tz-drag-handle { opacity: 0.5; }\n' +
'\n/* ---- New Filter Button ---- */\n' +
'.tz-new-filter { color: var(--tz-text-faint); font-style: italic; }\n' +
'.tz-new-filter:hover { color: var(--tz-text-muted); }\n' +
'\n/* ---- Context Menu ---- */\n' +
'.tz-context-menu {\n' +
'  position: fixed; z-index: 500;\n' +
'  background: var(--tz-bg-card); border: 1px solid var(--tz-border-strong);\n' +
'  border-radius: var(--tz-radius-sm); box-shadow: 0 8px 24px color-mix(in srgb, black 25%, transparent);\n' +
'  padding: 4px; min-width: 140px;\n' +
'}\n' +
'.tz-context-opt {\n' +
'  display: block; width: 100%; padding: 5px 10px; font-size: 12px;\n' +
'  border: none; background: transparent; color: var(--tz-text);\n' +
'  text-align: left; border-radius: var(--tz-radius-xs); cursor: pointer;\n' +
'  display: flex; align-items: center; gap: 8px;\n' +
'}\n' +
'.tz-context-opt:hover { background: var(--tz-border); }\n' +
'.tz-context-opt.danger { color: var(--tz-red); }\n' +
'.tz-context-opt.danger:hover { background: var(--tz-red-soft); }\n' +
'.tz-context-opt i { width: 14px; text-align: center; font-size: 11px; }\n' +
'.tz-context-sep { height: 1px; background: var(--tz-border); margin: 3px 4px; }\n' +
'.tz-filter-item.saved.is-dragging { opacity: 0.35; }\n' +
'.tz-filter-item.saved.drag-over-top {\n' +
'  box-shadow: 0 -2px 0 0 var(--tz-accent) inset;\n' +
'}\n' +
'.tz-filter-item.saved.drag-over-bottom {\n' +
'  box-shadow: 0 2px 0 0 var(--tz-accent) inset;\n' +
'}\n' +
'\n/* ---- Main Content ---- */\n' +
'.tz-main { flex: 1; display: flex; flex-direction: column; overflow: hidden; min-width: 0; }\n' +
'.tz-header {\n' +
'  display: flex; align-items: center; gap: 8px;\n' +
'  padding: 10px 14px; border-bottom: 1px solid var(--tz-border);\n' +
'  background: color-mix(in srgb, var(--tz-bg) 92%, transparent);\n' +
'  backdrop-filter: blur(16px); -webkit-backdrop-filter: blur(16px);\n' +
'}\n' +
'.tz-search-wrap {\n' +
'  flex: 1; display: flex; align-items: center; gap: 8px;\n' +
'  padding: 0 10px; height: 32px;\n' +
'  background: var(--tz-bg-card); border: 1px solid var(--tz-border-strong);\n' +
'  border-radius: var(--tz-radius-sm);\n' +
'  transition: border-color 0.15s ease;\n' +
'}\n' +
'.tz-search-wrap:focus-within { border-color: var(--tz-accent); }\n' +
'.tz-search-icon { font-size: 12px; color: var(--tz-text-faint); flex-shrink: 0; }\n' +
'.tz-search-input {\n' +
'  flex: 1; border: none; background: transparent; color: var(--tz-text);\n' +
'  font-size: 12px; font-family: "SF Mono", "Fira Code", monospace;\n' +
'  outline: none;\n' +
'}\n' +
'.tz-search-input::placeholder { color: var(--tz-text-faint); font-family: inherit; }\n' +
'.tz-search-clear {\n' +
'  width: 18px; height: 18px; display: flex; align-items: center; justify-content: center;\n' +
'  border-radius: 50%; border: none; background: transparent;\n' +
'  color: var(--tz-text-faint); cursor: pointer; font-size: 10px; flex-shrink: 0;\n' +
'}\n' +
'.tz-search-clear:hover { background: var(--tz-border); color: var(--tz-text); }\n' +
'.tz-header-actions { display: flex; gap: 6px; flex-shrink: 0; }\n' +
'.tz-save-area { flex-shrink: 0; position: relative; }\n' +
'.tz-btn {\n' +
'  display: inline-flex; align-items: center; gap: 5px;\n' +
'  padding: 5px 10px; font-size: 12px; font-weight: 500;\n' +
'  border-radius: var(--tz-radius-sm);\n' +
'  border: 1px solid var(--tz-border-strong);\n' +
'  background: var(--tz-bg-card); color: var(--tz-text);\n' +
'  cursor: pointer; transition: all 0.15s ease; white-space: nowrap;\n' +
'}\n' +
'.tz-btn:hover { background: var(--tz-bg-elevated); border-color: color-mix(in srgb, var(--tz-text) 25%, transparent); }\n' +
'.tz-btn:active { transform: scale(0.97); }\n' +
'.tz-btn i { font-size: 11px; }\n' +
'\n/* ---- Toolbar ---- */\n' +
'.tz-toolbar {\n' +
'  display: flex; align-items: center; gap: 4px;\n' +
'  padding: 6px 14px; border-bottom: 1px solid var(--tz-border);\n' +
'  overflow-x: auto;\n' +
'}\n' +
'.tz-toolbar-label { font-size: 11px; color: var(--tz-text-faint); font-weight: 600; margin-right: 4px; white-space: nowrap; }\n' +
'.tz-group-btn {\n' +
'  padding: 3px 9px; font-size: 11px; font-weight: 500;\n' +
'  border-radius: 100px; border: none; background: transparent;\n' +
'  color: var(--tz-text-muted); cursor: pointer;\n' +
'  transition: all 0.12s ease; white-space: nowrap;\n' +
'}\n' +
'.tz-group-btn:hover { background: var(--tz-border); color: var(--tz-text); }\n' +
'.tz-group-btn.active { background: var(--tz-accent-soft); color: var(--tz-accent); font-weight: 600; }\n' +
'.tz-toolbar-count {\n' +
'  margin-left: auto; font-size: 11px; color: var(--tz-text-faint);\n' +
'  font-variant-numeric: tabular-nums; white-space: nowrap;\n' +
'}\n' +
'\n/* ---- Body / Task List ---- */\n' +
'.tz-body { flex: 1; overflow-y: auto; padding: 12px 14px 40px; }\n' +
'.tz-group { margin-bottom: 16px; }\n' +
'.tz-group-header {\n' +
'  display: flex; align-items: center; gap: 8px;\n' +
'  padding: 6px 0 4px; margin-bottom: 4px;\n' +
'  border-bottom: 1px solid var(--tz-border);\n' +
'}\n' +
'.tz-group-label {\n' +
'  font-size: 12px; font-weight: 700; color: var(--tz-text-muted);\n' +
'  overflow: hidden; text-overflow: ellipsis; white-space: nowrap;\n' +
'}\n' +
'.tz-group-count {\n' +
'  display: inline-flex; align-items: center; justify-content: center;\n' +
'  min-width: 18px; height: 18px; padding: 0 5px;\n' +
'  font-size: 10px; font-weight: 700; border-radius: 100px;\n' +
'  background: var(--tz-border); color: var(--tz-text-muted);\n' +
'  font-variant-numeric: tabular-nums;\n' +
'}\n' +
'.tz-task-list { display: flex; flex-direction: column; }\n' +
'\n/* ---- Task Row ---- */\n' +
'.tz-task {\n' +
'  display: flex; align-items: flex-start; gap: 6px;\n' +
'  padding: 5px 6px; border-radius: var(--tz-radius-xs);\n' +
'  transition: background 0.1s ease; position: relative;\n' +
'}\n' +
'.tz-task:hover { background: var(--tz-border); }\n' +
'.tz-task.indent-1 { padding-left: 22px; }\n' +
'.tz-task.indent-2 { padding-left: 38px; }\n' +
'.tz-task.indent-3 { padding-left: 54px; }\n' +
'.tz-task-cb {\n' +
'  width: 18px; height: 18px; flex-shrink: 0; margin-top: 1px;\n' +
'  display: flex; align-items: center; justify-content: center;\n' +
'  border-radius: 50%; cursor: pointer; font-size: 14px; transition: all 0.15s ease;\n' +
'}\n' +
'.tz-task-cb.open { color: var(--tz-text-faint); }\n' +
'.tz-task-cb.open:hover { color: var(--tz-green); }\n' +
'.tz-task-cb.done { color: var(--tz-green); }\n' +
'.tz-task-cb.cancelled { color: var(--tz-text-faint); }\n' +
'.tz-task-content {\n' +
'  flex: 1; min-width: 0; font-size: 15px; line-height: 1.5;\n' +
'  word-break: break-word;\n' +
'}\n' +
'.tz-task.is-done .tz-task-content { text-decoration: line-through; color: var(--tz-text-faint); }\n' +
'.tz-task.is-cancelled .tz-task-content { text-decoration: line-through; color: var(--tz-text-faint); }\n' +
'.tz-task-pri {\n' +
'  display: inline-flex; align-items: center; justify-content: center;\n' +
'  padding: 0 4px; height: 16px; border-radius: 3px;\n' +
'  font-size: 9px; font-weight: 800; flex-shrink: 0; cursor: pointer;\n' +
'  margin-top: 2px; transition: all 0.15s ease;\n' +
'}\n' +
priCSS('tz-task-pri') +
'\n/* ---- Task Meta Badges ---- */\n' +
'.tz-task-meta {\n' +
'  display: flex; align-items: center; gap: 4px; flex-shrink: 0;\n' +
'  flex-wrap: wrap; margin-top: 1px;\n' +
'}\n' +
'.tz-task-badge {\n' +
'  display: inline-flex; align-items: center; gap: 3px;\n' +
'  padding: 0 6px; height: 18px; border-radius: 3px;\n' +
'  font-size: 10px; font-weight: 500; white-space: nowrap;\n' +
'  background: var(--tz-border); color: var(--tz-text-muted);\n' +
'  cursor: default; transition: all 0.1s ease;\n' +
'}\n' +
'.tz-task-badge i { font-size: 9px; }\n' +
'.tz-task-badge.note { cursor: pointer; }\n' +
'.tz-task-badge.note:hover { background: var(--tz-accent-soft); color: var(--tz-accent); }\n' +
'.tz-task-badge.date { cursor: pointer; }\n' +
'.tz-task-badge.date:hover { background: var(--tz-blue-soft); color: var(--tz-blue); }\n' +
'.tz-task-badge.date.overdue { background: var(--tz-red-soft); color: var(--tz-red); }\n' +
'.tz-task-badge.date.overdue:hover { background: var(--tz-red-soft); color: var(--tz-red); filter: brightness(1.15); }\n' +
'.tz-task-badge.date.today { background: var(--tz-orange-soft); color: var(--tz-orange); }\n' +
'.tz-task-badge.date.today:hover { background: var(--tz-orange-soft); color: var(--tz-orange); filter: brightness(1.15); }\n' +
'.tz-task-badge.date.this-week { background: var(--tz-yellow-soft); color: var(--tz-yellow); }\n' +
'.tz-task-badge.date.this-week:hover { background: var(--tz-yellow-soft); color: var(--tz-yellow); filter: brightness(1.15); }\n' +
'.tz-task-badge.due { background: var(--tz-yellow-soft); color: var(--tz-yellow); }\n' +
'.tz-task-badge.due.overdue { background: var(--tz-red-soft); color: var(--tz-red); }\n' +
'.tz-task-badge.folder { font-size: 10px; }\n' +
'\n/* ---- Markdown in Task Content ---- */\n' +
'.tz-task-content strong { font-weight: 700; }\n' +
'.tz-task-content em { font-style: italic; }\n' +
'.tz-task-content del { text-decoration: line-through; color: var(--tz-text-muted); }\n' +
'.tz-md-code {\n' +
'  font-family: "SF Mono", "Fira Code", "Menlo", monospace; font-size: 11px;\n' +
'  padding: 1px 4px; border-radius: 3px;\n' +
'  background: var(--tz-border); color: var(--tz-text);\n' +
'}\n' +
'.tz-md-link {\n' +
'  color: var(--tz-blue); text-decoration: none; cursor: pointer;\n' +
'}\n' +
'.tz-md-link:hover { text-decoration: underline; }\n' +
'.tz-md-highlight {\n' +
'  background: var(--tz-yellow-soft); color: var(--tz-yellow);\n' +
'  padding: 0 2px; border-radius: 2px;\n' +
'}\n' +
'.tz-tag, .tz-mention {\n' +
'  color: var(--tz-orange); font-weight: 600;\n' +
'}\n' +
'\n/* ---- Task Hover Actions ---- */\n' +
'.tz-task-acts {\n' +
'  display: none; align-items: center; gap: 2px; flex-shrink: 0; margin-top: 1px;\n' +
'}\n' +
'.tz-task:hover .tz-task-acts { display: flex; }\n' +
'.tz-task-act {\n' +
'  width: 20px; height: 20px; display: flex; align-items: center; justify-content: center;\n' +
'  border-radius: 3px; border: none; background: transparent;\n' +
'  color: var(--tz-text-faint); cursor: pointer; font-size: 10px; transition: all 0.1s ease;\n' +
'}\n' +
'.tz-task-act:hover { background: var(--tz-border-strong); color: var(--tz-text); }\n' +
'.tz-task-act.cancel:hover { color: var(--tz-red); }\n' +
'\n/* ---- Schedule Picker ---- */\n' +
'.tz-sched-picker {\n' +
'  position: fixed; z-index: 500;\n' +
'  background: var(--tz-bg-card); border: 1px solid var(--tz-border-strong);\n' +
'  border-radius: var(--tz-radius-sm); box-shadow: 0 8px 24px color-mix(in srgb, black 25%, transparent);\n' +
'  padding: 4px; min-width: 150px;\n' +
'}\n' +
'.tz-sched-opt {\n' +
'  display: block; width: 100%; padding: 5px 10px; font-size: 12px;\n' +
'  border: none; background: transparent; color: var(--tz-text);\n' +
'  text-align: left; border-radius: var(--tz-radius-xs); cursor: pointer;\n' +
'}\n' +
'.tz-sched-opt:hover { background: var(--tz-border); }\n' +
'.tz-sched-opt.danger { color: var(--tz-red); }\n' +
'.tz-sched-date-input {\n' +
'  width: 100%; padding: 4px 8px; margin: 2px 0;\n' +
'  font-size: 12px; border: 1px solid var(--tz-border-strong);\n' +
'  border-radius: var(--tz-radius-xs); background: var(--tz-bg);\n' +
'  color: var(--tz-text); outline: none;\n' +
'}\n' +
'\n/* ---- Assign Picker ---- */\n' +
'.tz-assign-picker { min-width: 180px; }\n' +
'.tz-assign-search {\n' +
'  width: calc(100% - 16px); margin: 4px 8px 6px; padding: 5px 8px;\n' +
'  font-size: 13px; border: 1px solid var(--tz-border-strong);\n' +
'  border-radius: var(--tz-radius-xs); background: var(--tz-bg);\n' +
'  color: var(--tz-text); outline: none;\n' +
'}\n' +
'.tz-assign-search:focus { border-color: var(--tz-accent); }\n' +
'.tz-assign-options { max-height: 180px; overflow-y: auto; }\n' +
'.tz-sched-opt.new-mention { color: var(--tz-accent); font-style: italic; }\n' +
'\n/* ---- Save Filter Modal ---- */\n' +
'.tz-modal-overlay {\n' +
'  position: fixed; inset: 0; z-index: 600;\n' +
'  background: color-mix(in srgb, black 50%, transparent);\n' +
'  display: flex; align-items: center; justify-content: center;\n' +
'  backdrop-filter: blur(4px);\n' +
'}\n' +
'.tz-modal {\n' +
'  background: var(--tz-bg-card); border: 1px solid var(--tz-border-strong);\n' +
'  border-radius: var(--tz-radius); padding: 20px;\n' +
'  min-width: 300px; max-width: 400px;\n' +
'  box-shadow: 0 16px 48px color-mix(in srgb, black 30%, transparent);\n' +
'}\n' +
'.tz-modal-title { font-size: 14px; font-weight: 700; margin-bottom: 12px; }\n' +
'.tz-modal-input {\n' +
'  width: 100%; padding: 7px 10px; font-size: 13px;\n' +
'  border: 1px solid var(--tz-border-strong); border-radius: var(--tz-radius-sm);\n' +
'  background: var(--tz-bg); color: var(--tz-text); outline: none;\n' +
'  margin-bottom: 8px;\n' +
'}\n' +
'.tz-modal-input:focus { border-color: var(--tz-accent); }\n' +
'.tz-modal-query {\n' +
'  font-size: 11px; color: var(--tz-text-muted); margin-bottom: 14px;\n' +
'  padding: 6px 8px; background: var(--tz-border); border-radius: var(--tz-radius-xs);\n' +
'  font-family: "SF Mono", "Fira Code", monospace; word-break: break-all;\n' +
'}\n' +
'.tz-modal-actions { display: flex; justify-content: flex-end; gap: 8px; }\n' +
'.tz-modal-btn {\n' +
'  padding: 6px 14px; font-size: 12px; font-weight: 600;\n' +
'  border-radius: var(--tz-radius-sm); border: 1px solid var(--tz-border-strong);\n' +
'  background: var(--tz-bg-card); color: var(--tz-text);\n' +
'  cursor: pointer; transition: all 0.15s ease;\n' +
'}\n' +
'.tz-modal-btn:hover { background: var(--tz-bg-elevated); }\n' +
'.tz-modal-btn.primary {\n' +
'  background: var(--tz-accent); color: #fff; border-color: var(--tz-accent);\n' +
'}\n' +
'.tz-modal-btn.primary:hover { filter: brightness(1.1); }\n' +
'\n/* ---- Save Area & Dropdown ---- */\n' +
'.tz-save-dropdown {\n' +
'  position: absolute; top: 100%; right: 0; z-index: 200;\n' +
'  min-width: 180px; padding: 4px;\n' +
'  background: var(--tz-bg-card); border: 1px solid var(--tz-border-strong);\n' +
'  border-radius: var(--tz-radius-sm); margin-top: 4px;\n' +
'  box-shadow: 0 8px 24px color-mix(in srgb, black 25%, transparent);\n' +
'}\n' +
'.tz-save-dropdown-opt {\n' +
'  display: flex; align-items: center; gap: 8px; width: 100%;\n' +
'  padding: 7px 10px; font-size: 12px; font-weight: 500;\n' +
'  color: var(--tz-text); background: transparent;\n' +
'  border: none; border-radius: var(--tz-radius-xs);\n' +
'  cursor: pointer; text-align: left; white-space: nowrap;\n' +
'}\n' +
'.tz-save-dropdown-opt:hover { background: var(--tz-border); }\n' +
'.tz-save-dropdown-opt i { color: var(--tz-text-muted); font-size: 11px; width: 14px; text-align: center; }\n' +
'\n/* ---- Empty State ---- */\n' +
'.tz-empty {\n' +
'  display: flex; flex-direction: column; align-items: center;\n' +
'  justify-content: center; padding: 60px 20px; text-align: center;\n' +
'}\n' +
'.tz-empty-icon { font-size: 36px; color: var(--tz-text-faint); margin-bottom: 12px; }\n' +
'.tz-empty-title { font-size: 15px; font-weight: 600; margin-bottom: 4px; }\n' +
'.tz-empty-desc { font-size: 12px; color: var(--tz-text-muted); }\n' +
'\n/* ---- Toast ---- */\n' +
'.tz-toast {\n' +
'  position: fixed; bottom: 20px; right: 20px;\n' +
'  padding: 10px 16px; font-size: 12px; font-weight: 500;\n' +
'  background: var(--tz-green); color: #fff;\n' +
'  border-radius: var(--tz-radius-sm);\n' +
'  box-shadow: 0 4px 16px color-mix(in srgb, black 20%, transparent);\n' +
'  z-index: 1000;\n' +
'  animation: tzToastIn 0.3s ease, tzToastOut 0.3s ease 2.5s forwards;\n' +
'}\n' +
'@keyframes tzToastIn { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: translateY(0); } }\n' +
'@keyframes tzToastOut { from { opacity: 1; } to { opacity: 0; transform: translateY(8px); } }\n' +
'\n/* ---- Tooltips ---- */\n' +
'[data-tooltip] { position: relative; }\n' +
'[data-tooltip]:hover::after {\n' +
'  content: attr(data-tooltip); position: absolute; top: calc(100% + 6px); left: 50%; transform: translateX(-50%);\n' +
'  padding: 4px 8px; font-size: 11px; font-weight: 500;\n' +
'  white-space: nowrap; background: var(--tz-bg-elevated); color: var(--tz-text);\n' +
'  border: 1px solid var(--tz-border-strong); border-radius: var(--tz-radius-xs);\n' +
'  z-index: 500; pointer-events: none;\n' +
'}\n' +
'/* Right-edge tooltips: align right instead of center */\n' +
'.tz-save-area [data-tooltip]:hover::after { left: auto; right: 0; transform: none; }\n' +
'\n/* ---- Scrollbar ---- */\n' +
'::-webkit-scrollbar { width: 6px; }\n' +
'::-webkit-scrollbar-track { background: transparent; }\n' +
'::-webkit-scrollbar-thumb { background: var(--tz-border-strong); border-radius: 100px; }\n' +
'::-webkit-scrollbar-thumb:hover { background: color-mix(in srgb, var(--tz-text) 30%, transparent); }\n' +
'\n/* ---- Animations ---- */\n' +
'@keyframes tzFadeIn { from { opacity: 0; transform: translateY(3px); } to { opacity: 1; transform: translateY(0); } }\n' +
'.tz-task { animation: tzFadeIn 0.2s ease both; }\n' +
'.tz-group:nth-child(1) .tz-task { animation-delay: 0.01s; }\n' +
'.tz-group:nth-child(2) .tz-task { animation-delay: 0.03s; }\n' +
'.tz-group:nth-child(3) .tz-task { animation-delay: 0.05s; }\n' +
'\n/* ---- Sidebar Toggle Button (hidden on desktop) ---- */\n' +
'.tz-sidebar-toggle {\n' +
'  display: none; width: 32px; height: 32px; flex-shrink: 0;\n' +
'  align-items: center; justify-content: center;\n' +
'  border-radius: var(--tz-radius-sm); border: 1px solid var(--tz-border-strong);\n' +
'  background: var(--tz-bg-card); color: var(--tz-text-muted);\n' +
'  cursor: pointer; font-size: 13px;\n' +
'}\n' +
'.tz-sidebar-toggle:hover { background: var(--tz-bg-elevated); color: var(--tz-text); }\n' +
'\n/* ---- Mobile Overlay Backdrop ---- */\n' +
'.tz-sidebar-backdrop {\n' +
'  display: none; position: fixed; inset: 0; z-index: 90;\n' +
'  background: color-mix(in srgb, black 40%, transparent);\n' +
'}\n' +
'\n/* ---- Responsive: Narrow / Mobile ---- */\n' +
'@media (max-width: 600px) {\n' +
'  .tz-sidebar-toggle { display: flex; }\n' +
'  .tz-sidebar {\n' +
'    position: fixed; left: 0; top: 0; bottom: 0; z-index: 100;\n' +
'    width: 220px; transform: translateX(-100%);\n' +
'    transition: transform 0.25s cubic-bezier(0.22, 1, 0.36, 1);\n' +
'    box-shadow: none;\n' +
'  }\n' +
'  .tz-sidebar.open {\n' +
'    transform: translateX(0);\n' +
'    box-shadow: 4px 0 24px color-mix(in srgb, black 25%, transparent);\n' +
'  }\n' +
'  .tz-sidebar-backdrop.open { display: block; }\n' +
'  .tz-layout { flex-direction: column; }\n' +
'  .tz-main { width: 100%; }\n' +
'  .tz-header { flex-wrap: wrap; gap: 6px; padding: 8px 10px; }\n' +
'  .tz-search-wrap { min-width: 0; order: 2; flex-basis: 100%; height: 36px; }\n' +
'  .tz-sidebar-toggle { order: 0; }\n' +
'  .tz-save-area { order: 1; margin-left: auto; }\n' +
'  .tz-toolbar { padding: 6px 10px; gap: 2px; }\n' +
'  .tz-group-btn { padding: 4px 7px; font-size: 10px; }\n' +
'  .tz-toolbar-label { font-size: 10px; }\n' +
'  .tz-toolbar-count { font-size: 10px; }\n' +
'  .tz-body { padding: 8px 8px 40px; }\n' +
'  .tz-task {\n' +
'    flex-wrap: wrap; gap: 4px; padding: 8px 6px;\n' +
'  }\n' +
'  .tz-task-content {\n' +
'    flex: 1 1 0; min-width: 0; width: 0;\n' +
'  }\n' +
'  .tz-task-meta {\n' +
'    flex-basis: 100%; padding-left: 24px; margin-top: 2px;\n' +
'  }\n' +
'  .tz-task-acts {\n' +
'    display: flex; flex-shrink: 0; margin-left: auto;\n' +
'  }\n' +
'  .tz-task:hover .tz-task-acts { display: flex; }\n' +
'  .tz-btn span { display: none; }\n' +
'  .tz-sched-picker { left: 10px !important; right: 10px; min-width: auto; }\n' +
'  .tz-modal { min-width: auto; margin: 16px; max-width: calc(100vw - 32px); }\n' +
'  [data-tooltip]:hover::after { display: none; }\n' +
'}\n' +
'';
}

// ============================================
// NOTE LOOKUP — supports both project and calendar notes
// ============================================

/**
 * Find a note by filename, checking project notes first then calendar notes.
 */
function findNoteByFilename(filename) {
  // Try project note first
  var note = DataStore.projectNoteByFilename(filename);
  if (note) return note;

  // Try calendar note by filename
  var calNotes = DataStore.calendarNotes;
  for (var i = 0; i < calNotes.length; i++) {
    if (calNotes[i].filename === filename) return calNotes[i];
  }

  // Try calendarNoteByDateString for daily notes (YYYYMMDD → YYYY-MM-DD)
  var dailyMatch = filename.replace(/\.(md|txt)$/, '').match(/^(\d{4})(\d{2})(\d{2})$/);
  if (dailyMatch) {
    var dateStr = dailyMatch[1] + '-' + dailyMatch[2] + '-' + dailyMatch[3];
    try {
      var calNote = DataStore.calendarNoteByDateString(dateStr);
      if (calNote) return calNote;
    } catch (e) {}
  }

  return null;
}

/**
 * Check if a filename refers to a calendar note
 */
function isCalendarFilename(filename) {
  var base = filename.replace(/\.(md|txt)$/, '');
  return /^\d{8}$/.test(base) || /^\d{4}-W\d{2}$/.test(base) ||
         /^\d{4}-\d{2}$/.test(base) || /^\d{4}-Q[1-4]$/.test(base) ||
         /^\d{4}$/.test(base);
}

// ============================================
// TASK MUTATIONS
// ============================================

function toggleTaskComplete(filename, lineIndex) {
  var note = findNoteByFilename(filename);
  if (!note) return null;
  var para = note.paragraphs[lineIndex];
  if (!para) return null;

  // Detect if this is a checklist item from rawContent (+ marker)
  var rawLine = (para.rawContent || '').trimStart();
  var isChecklist = rawLine.startsWith('+');

  if (para.type === 'done' || para.type === 'checklistDone') {
    // Restore to the correct open type
    para.type = isChecklist ? 'checklist' : 'open';
    para.content = (para.content || '').replace(/\s*@done\([^)]*\)/, '');
  } else {
    // Complete — append @done(date) for Routine compatibility
    para.type = isChecklist ? 'checklistDone' : 'done';
    var now = new Date();
    var h = now.getHours(); var mi = String(now.getMinutes()).padStart(2, '0');
    var ampm = h >= 12 ? 'PM' : 'AM'; var h12 = h % 12; if (h12 === 0) h12 = 12;
    var doneStr = '@done(' + now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0') + '-' + String(now.getDate()).padStart(2, '0') + ' ' + String(h12).padStart(2, '0') + ':' + mi + ' ' + ampm + ')';
    para.content = (para.content || '').trimEnd() + ' ' + doneStr;
  }
  note.updateParagraph(para);
  return { lineIndex: lineIndex, newType: para.type };
}

function toggleTaskCancel(filename, lineIndex) {
  var note = findNoteByFilename(filename);
  if (!note) return null;
  var para = note.paragraphs[lineIndex];
  if (!para) return null;

  // Detect if this is a checklist item from rawContent (+ marker)
  var rawLine = (para.rawContent || '').trimStart();
  var isChecklist = rawLine.startsWith('+');

  if (para.type === 'cancelled' || para.type === 'checklistCancelled') {
    para.type = isChecklist ? 'checklist' : 'open';
  } else {
    para.type = isChecklist ? 'checklistCancelled' : 'cancelled';
  }
  note.updateParagraph(para);
  return { lineIndex: lineIndex, newType: para.type };
}

function cycleTaskPriority(filename, lineIndex) {
  var note = findNoteByFilename(filename);
  if (!note) return null;
  var para = note.paragraphs[lineIndex];
  if (!para) return null;

  var content = para.content || '';
  var currentPri = 0;
  if (content.startsWith('!!! ')) { currentPri = 3; content = content.slice(4); }
  else if (content.startsWith('!! ')) { currentPri = 2; content = content.slice(3); }
  else if (content.startsWith('! ')) { currentPri = 1; content = content.slice(2); }

  var nextPri = (currentPri + 1) % 4;
  var prefix = nextPri === 3 ? '!!! ' : nextPri === 2 ? '!! ' : nextPri === 1 ? '! ' : '';
  para.content = prefix + content;
  note.updateParagraph(para);
  return { lineIndex: lineIndex, newPriority: nextPri };
}

function assignPerson(filename, lineIndex, mention) {
  var note = findNoteByFilename(filename);
  if (!note) return null;
  var para = note.paragraphs[lineIndex];
  if (!para) return null;

  // Ensure mention starts with @
  if (!mention.startsWith('@')) mention = '@' + mention;

  // Append mention to content
  para.content = (para.content || '').trimEnd() + ' ' + mention;

  // Convert task to checklist: open → checklist, done → checklistDone, cancelled → checklistCancelled
  if (para.type === 'open') para.type = 'checklist';
  else if (para.type === 'done') para.type = 'checklistDone';
  else if (para.type === 'cancelled') para.type = 'checklistCancelled';

  note.updateParagraph(para);
  return { lineIndex: lineIndex };
}

function scheduleTask(filename, lineIndex, dateStr) {
  var note = findNoteByFilename(filename);
  if (!note) return null;
  var para = note.paragraphs[lineIndex];
  if (!para) return null;

  var content = para.content || '';
  content = content.replace(/\s*>\d{4}-\d{2}-\d{2}(\s+\d{1,2}:\d{2}\s*(AM|PM)(\s*-\s*\d{1,2}:\d{2}\s*(AM|PM))?)?/gi, '');
  content = content.replace(/\s*>\d{4}-W\d{2}/g, '');
  content = content.replace(/\s*>today/g, '');

  if (dateStr) {
    content = content.trimEnd() + ' >' + dateStr;
  }

  para.content = content;
  note.updateParagraph(para);
  return { lineIndex: lineIndex, scheduledDate: dateStr };
}

// ============================================
// PLUGIN COMMANDS
// ============================================

async function showTaskZoom(activeQuery, activeFilterId, groupBy) {
  try {
    CommandBar.showLoading(true, 'Scanning tasks...');
    await CommandBar.onAsyncThread();

    var config = getSettings();
    var prefs = getUserPrefs();

    // Restore last session state if no explicit args
    var filterId = activeFilterId || prefs.lastFilterId || '__overdue';
    var query = activeQuery || prefs.lastQuery || 'open overdue';
    // Restore per-filter groupBy
    var group = groupBy || prefs.groupByMap[filterId] || config.defaultGroupBy || 'note';

    // Persist current choice
    saveUserPrefs(filterId, query, group);

    var bodyContent = buildDashboardHTML(config, query, filterId, group);
    var fullHTML = buildFullHTML(bodyContent);

    await CommandBar.onMainThread();
    CommandBar.showLoading(false);

    var winOptions = {
      customId: WINDOW_ID,
      savedFilename: '../../asktru.TaskZoom/task_zoom.html',
      shouldFocus: true,
      reuseUsersWindowRect: true,
      headerBGColor: 'transparent',
      autoTopPadding: true,
      showReloadButton: true,
      reloadPluginID: PLUGIN_ID,
      reloadCommandName: 'Task Zoom',
      icon: 'fa-magnifying-glass-plus',
      iconColor: '#F59E0B',
    };

    var result = await HTMLView.showInMainWindow(fullHTML, 'Task Zoom', winOptions);
    if (!result || !result.success) {
      console.log('TaskZoom: showInMainWindow failed, falling back to floating window');
      await HTMLView.showWindowWithOptions(fullHTML, 'Task Zoom', winOptions);
    }
  } catch (err) {
    CommandBar.showLoading(false);
    console.log('TaskZoom error: ' + String(err));
  }
}

async function refreshTaskZoom() {
  invalidateTaskCache();
  await showTaskZoom();
}

/**
 * Handle messages from the HTML window
 */
async function onMessageFromHTMLView(actionType, data) {
  try {
    var parsedData = typeof data === 'string' ? JSON.parse(data) : data;

    switch (actionType) {
      case 'runFilter': {
        var rfConfig = getSettings();
        var rfFilterId = parsedData.filterId || '__overdue';
        var rfQuery = parsedData.query || 'open overdue';
        var rfPrefs = getUserPrefs();
        var rfGroup = parsedData.groupBy || rfPrefs.groupByMap[rfFilterId] || rfConfig.defaultGroupBy || 'note';

        saveUserPrefs(rfFilterId, rfQuery, rfGroup);

        // Look up the original query for save-button detection
        var rfBuiltinQueries = {
          '__overdue': 'open overdue', '__high': 'open p1 | open p2 | open p3',
          '__today': 'open today', '__thisweek': 'open this week',
          '__nodate': 'open no date', '__all': 'open',
        };
        var rfOrigQuery = rfBuiltinQueries[rfFilterId] || '';
        if (!rfOrigQuery) {
          for (var rfs = 0; rfs < rfConfig.savedFilters.length; rfs++) {
            if (rfConfig.savedFilters[rfs].id === rfFilterId) { rfOrigQuery = rfConfig.savedFilters[rfs].query; break; }
          }
        }

        // Check filter result cache (caches filtered tasks, not HTML — grouping is cheap)
        var rfCacheKey = getFilterCacheKey(rfQuery);
        var rfFiltered;
        var rfAllTasks = getCachedTasks();

        if (_filterResultCache[rfCacheKey]) {
          rfFiltered = _filterResultCache[rfCacheKey];
        } else {
          var rfQueryLower = (rfQuery || '').toLowerCase();
          var rfHasKind = /\b(checklist|checklists|task|tasks)\b/.test(rfQueryLower);
          var rfFilter = parseQuery(rfQuery);
          rfFiltered = [];
          for (var rfi = 0; rfi < rfAllTasks.length; rfi++) {
            if (!rfHasKind && rfAllTasks[rfi].taskKind !== 'task') continue;
            if (evaluateFilter(rfAllTasks[rfi], rfFilter)) rfFiltered.push(rfAllTasks[rfi]);
          }
          _filterResultCache[rfCacheKey] = rfFiltered;
        }

        // Build HTML (grouping + rendering is fast)
        var rfBodyHTML = '';
        if (rfFiltered.length === 0) {
          rfBodyHTML = '<div class="tz-empty"><div class="tz-empty-icon"><i class="fa-solid fa-filter-circle-xmark"></i></div><div class="tz-empty-title">No tasks match this filter</div><div class="tz-empty-desc">Try adjusting your search query or filter settings.</div></div>';
        } else {
          var rfGroups = groupTasks(rfFiltered, rfGroup);
          for (var rfgi = 0; rfgi < rfGroups.length; rfgi++) {
            var rfGrp = rfGroups[rfgi];
            rfBodyHTML += '<div class="tz-group"><div class="tz-group-header"><span class="tz-group-label">' + esc(rfGrp.label) + '</span><span class="tz-group-count">' + rfGrp.tasks.length + '</span></div><div class="tz-task-list">';
            for (var rfti = 0; rfti < rfGrp.tasks.length; rfti++) {
              rfBodyHTML += buildTaskRow(rfGrp.tasks[rfti], rfGroup);
            }
            rfBodyHTML += '</div></div>';
          }
        }
        var rfCount = rfFiltered.length + ' tasks / ' + rfAllTasks.length + ' scanned';

        await sendToHTMLWindow('asktru.TaskZoom.dashboard', 'FILTER_RESULTS', {
          bodyHTML: rfBodyHTML,
          taskCount: rfCount,
          query: rfQuery,
          filterId: rfFilterId,
          groupBy: rfGroup,
          originalQuery: rfOrigQuery,
        });
        break;
      }

      case 'toggleTaskComplete': {
        invalidateTaskCache();
        var fn1 = decSafe(parsedData.encodedFilename);
        var myWinId = 'asktru.TaskZoom.dashboard';
        // Check for @repeat before toggling
        var tc1Note = findNoteByFilename(fn1);
        var tc1Para = tc1Note ? tc1Note.paragraphs[parsedData.lineIndex] : null;
        var tc1HasRepeat = tc1Para && (tc1Para.content || '').indexOf('@repeat') >= 0;
        var tc1WasOpen = tc1Para && (tc1Para.type === 'open' || tc1Para.type === 'checklist');
        var result1 = toggleTaskComplete(fn1, parsedData.lineIndex);
        if (result1) {
          var rawType1 = result1.newType;
          var isCL1 = rawType1 === 'checklistDone' || rawType1 === 'checklist' || rawType1 === 'checklistCancelled';
          var uiType1 = rawType1;
          if (uiType1 === 'checklistDone') uiType1 = 'done';
          else if (uiType1 === 'checklist') uiType1 = 'open';
          else if (uiType1 === 'checklistCancelled') uiType1 = 'cancelled';
          await sendToHTMLWindow(myWinId, 'TASK_UPDATED', {
            encodedFilename: parsedData.encodedFilename,
            lineIndex: result1.lineIndex,
            newType: uiType1,
            isChecklist: isCL1,
          });
          // Invoke Routine for repeating tasks
          if (tc1HasRepeat && tc1WasOpen && (rawType1 === 'done' || rawType1 === 'checklistDone')) {
            try {
              await DataStore.invokePluginCommandByName('generate repeats', 'asktru.Routine', [fn1]);
            } catch (e) { console.log('TaskZoom: Routine not available: ' + String(e)); }
          }
        }
        break;
      }

      case 'toggleTaskCancel': {
        invalidateTaskCache();
        var fn2 = decSafe(parsedData.encodedFilename);
        var myWinId2 = 'asktru.TaskZoom.dashboard';
        var tc2Note = findNoteByFilename(fn2);
        var tc2Para = tc2Note ? tc2Note.paragraphs[parsedData.lineIndex] : null;
        var tc2HasRepeat = tc2Para && (tc2Para.content || '').indexOf('@repeat') >= 0;
        var tc2WasOpen = tc2Para && (tc2Para.type === 'open' || tc2Para.type === 'checklist');
        var result2 = toggleTaskCancel(fn2, parsedData.lineIndex);
        if (result2) {
          var rawType2 = result2.newType;
          var isCL2 = rawType2 === 'checklistDone' || rawType2 === 'checklist' || rawType2 === 'checklistCancelled';
          var uiType2 = rawType2;
          if (uiType2 === 'checklistDone') uiType2 = 'done';
          else if (uiType2 === 'checklist') uiType2 = 'open';
          else if (uiType2 === 'checklistCancelled') uiType2 = 'cancelled';
          await sendToHTMLWindow(myWinId2, 'TASK_UPDATED', {
            encodedFilename: parsedData.encodedFilename,
            lineIndex: result2.lineIndex,
            newType: uiType2,
            isChecklist: isCL2,
          });
          if (tc2HasRepeat && tc2WasOpen && (rawType2 === 'cancelled' || rawType2 === 'checklistCancelled')) {
            try {
              await DataStore.invokePluginCommandByName('generate repeats', 'asktru.Routine', [fn2]);
            } catch (e) { console.log('TaskZoom: Routine not available: ' + String(e)); }
          }
        }
        break;
      }

      case 'cycleTaskPriority': {
        invalidateTaskCache();
        var fn3 = decSafe(parsedData.encodedFilename);
        var result3 = cycleTaskPriority(fn3, parsedData.lineIndex);
        if (result3) {
          await sendToHTMLWindow(WINDOW_ID, 'TASK_PRIORITY_CHANGED', {
            encodedFilename: parsedData.encodedFilename,
            lineIndex: result3.lineIndex,
            newPriority: result3.newPriority,
          });
        }
        break;
      }

      case 'scheduleTask': {
        invalidateTaskCache();
        var fn4 = decSafe(parsedData.encodedFilename);
        var result4 = scheduleTask(fn4, parsedData.lineIndex, parsedData.dateStr);
        if (result4) {
          await sendToHTMLWindow(WINDOW_ID, 'TASK_SCHEDULED', {
            encodedFilename: parsedData.encodedFilename,
            lineIndex: result4.lineIndex,
            scheduledDate: result4.scheduledDate,
          });
        }
        break;
      }

      case 'openNote': {
        var fn5 = decSafe(parsedData.encodedFilename);
        await CommandBar.onMainThread();
        if (isCalendarFilename(fn5)) {
          var dailyM = fn5.replace(/\.(md|txt)$/, '').match(/^(\d{4})(\d{2})(\d{2})$/);
          if (dailyM) {
            var dateStr5 = dailyM[1] + '-' + dailyM[2] + '-' + dailyM[3];
            NotePlan.openURL('noteplan://x-callback-url/openNote?noteDate=' + encodeURIComponent(dateStr5) + '&splitView=yes&reuseSplitView=yes');
          } else {
            await Editor.openNoteByFilename(fn5);
          }
        } else {
          var note5 = findNoteByFilename(fn5);
          var title5 = note5 ? (note5.title || '') : '';
          if (title5) {
            NotePlan.openURL('noteplan://x-callback-url/openNote?noteTitle=' + encodeURIComponent(title5) + '&splitView=yes&reuseSplitView=yes');
          } else {
            await Editor.openNoteByFilename(fn5);
          }
        }
        break;
      }

      case 'saveFilter': {
        var cfg1 = getSettings();
        var newFilter = {
          id: 'f_' + Date.now(),
          name: parsedData.name,
          query: parsedData.query,
        };
        cfg1.savedFilters.push(newFilter);
        saveFilters(cfg1.savedFilters);
        await showTaskZoom(parsedData.query, newFilter.id, parsedData.groupBy);
        break;
      }

      case 'updateFilter': {
        var cfgUpd = getSettings();
        for (var ui = 0; ui < cfgUpd.savedFilters.length; ui++) {
          if (cfgUpd.savedFilters[ui].id === parsedData.filterId) {
            cfgUpd.savedFilters[ui].query = parsedData.query;
            break;
          }
        }
        saveFilters(cfgUpd.savedFilters);
        sendToHTMLWindow('SHOW_TOAST', { message: 'Filter updated' });
        await showTaskZoom(parsedData.query, parsedData.filterId, parsedData.groupBy);
        break;
      }

      case 'deleteFilter': {
        var cfg2 = getSettings();
        cfg2.savedFilters = cfg2.savedFilters.filter(function(f) { return f.id !== parsedData.filterId; });
        saveFilters(cfg2.savedFilters);
        await showTaskZoom(parsedData.currentQuery, '__overdue', parsedData.groupBy);
        break;
      }

      case 'renameFilter': {
        var cfg2r = getSettings();
        for (var ri2 = 0; ri2 < cfg2r.savedFilters.length; ri2++) {
          if (cfg2r.savedFilters[ri2].id === parsedData.filterId) {
            cfg2r.savedFilters[ri2].name = parsedData.newName;
            break;
          }
        }
        saveFilters(cfg2r.savedFilters);
        // Update the DOM inline via message
        await sendToHTMLWindow('asktru.TaskZoom.dashboard', 'FILTER_RENAMED', {
          filterId: parsedData.filterId,
          newName: parsedData.newName,
        });
        break;
      }

      case 'reorderFilters': {
        var cfg3 = getSettings();
        var orderedIds = parsedData.orderedIds; // array of filter IDs in new order
        if (orderedIds && orderedIds.length > 0) {
          var idToFilter = {};
          for (var ri = 0; ri < cfg3.savedFilters.length; ri++) {
            idToFilter[cfg3.savedFilters[ri].id] = cfg3.savedFilters[ri];
          }
          var reordered = [];
          for (var oi = 0; oi < orderedIds.length; oi++) {
            if (idToFilter[orderedIds[oi]]) reordered.push(idToFilter[orderedIds[oi]]);
          }
          // Append any filters not in the ordered list (safety net)
          for (var si = 0; si < cfg3.savedFilters.length; si++) {
            if (orderedIds.indexOf(cfg3.savedFilters[si].id) === -1) reordered.push(cfg3.savedFilters[si]);
          }
          saveFilters(reordered);
        }
        // No full refresh needed — the DOM was already reordered by the drag
        break;
      }

      case 'assignPerson': {
        var fn6 = decSafe(parsedData.encodedFilename);
        var result6 = assignPerson(fn6, parsedData.lineIndex, parsedData.mention);
        if (result6) {
          // Full refresh to show updated content
          await showTaskZoom(parsedData.query || 'open', parsedData.filterId, parsedData.groupBy);
        }
        break;
      }

      case 'refresh': {
        await showTaskZoom(parsedData.query, parsedData.filterId, parsedData.groupBy);
        break;
      }

      default:
        console.log('TaskZoom: unhandled action: ' + actionType);
    }
  } catch (err) {
    console.log('TaskZoom onMessage error: ' + String(err));
  }
}

/**
 * Send message to HTML window
 */
async function sendToHTMLWindow(windowId, type, data) {
  try {
    if (typeof HTMLView === 'undefined' || typeof HTMLView.runJavaScript !== 'function') {
      console.log('sendToHTMLWindow: HTMLView API not available');
      return;
    }
    var payload = {};
    var keys = Object.keys(data);
    for (var k = 0; k < keys.length; k++) payload[keys[k]] = data[keys[k]];
    payload.NPWindowID = windowId;

    var stringifiedPayload = JSON.stringify(payload);
    var doubleStringified = JSON.stringify(stringifiedPayload);
    var jsCode = '(function() { try { var pd = ' + doubleStringified + '; var p = JSON.parse(pd); window.postMessage({ type: "' + type + '", payload: p }, "*"); } catch(e) { console.error("sendToHTMLWindow error:", e); } })();';
    await HTMLView.runJavaScript(jsCode, windowId);
  } catch (err) {
    console.log('sendToHTMLWindow error: ' + String(err));
  }
}

// ============================================
// EXPORTS
// ============================================

globalThis.showTaskZoom = showTaskZoom;
globalThis.onMessageFromHTMLView = onMessageFromHTMLView;
globalThis.refreshTaskZoom = refreshTaskZoom;
