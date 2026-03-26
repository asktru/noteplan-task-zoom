/* ============================================
   asktru.TaskZoom — HTML Window Event Handlers
   Runs inside the HTML WebView window
   ============================================ */

// receivingPluginID is set in the inline script before the bridge loads

// ============================================
// STATE
// ============================================

var currentQuery = '';
var currentFilterId = '__all';
var currentGroupBy = 'note';
var originalQuery = ''; // tracks the query as loaded from the filter, to detect edits

// ============================================
// DATE HELPERS (client-side)
// ============================================

function todayStr() {
  var d = new Date();
  return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
}

function tomorrowStr() {
  var d = new Date();
  d.setDate(d.getDate() + 1);
  return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
}

function isoWeekStr(date) {
  var d = new Date(date.getTime());
  d.setHours(0,0,0,0);
  d.setDate(d.getDate() + 3 - (d.getDay() + 6) % 7);
  var week1 = new Date(d.getFullYear(), 0, 4);
  var weekNum = 1 + Math.round(((d - week1) / 86400000 - 3 + (week1.getDay() + 6) % 7) / 7);
  return d.getFullYear() + '-W' + String(weekNum).padStart(2, '0');
}

function thisWeekStr() {
  return isoWeekStr(new Date());
}

function nextWeekStr() {
  var d = new Date();
  d.setDate(d.getDate() + 7);
  return isoWeekStr(d);
}

// ============================================
// MESSAGE HANDLER FROM PLUGIN
// ============================================

function onMessageFromPlugin(type, data) {
  switch (type) {
    case 'SHOW_TOAST':
      showToast(data.message);
      break;
    case 'TASK_UPDATED':
      handleTaskUpdated(data);
      break;
    case 'TASK_PRIORITY_CHANGED':
      handlePriorityChanged(data);
      break;
    case 'TASK_SCHEDULED':
      handleTaskScheduled(data);
      break;
    case 'FULL_REFRESH':
      window.location.reload();
      break;
    default:
      console.log('TaskZoom: unknown message type', type);
  }
}

// ============================================
// TASK UPDATE HANDLERS
// ============================================

function handleTaskUpdated(data) {
  // Find ALL instances of this task (same task can appear in multiple groups)
  var taskEls = document.querySelectorAll('.tz-task[data-encoded-filename="' + data.encodedFilename + '"][data-line-index="' + data.lineIndex + '"]');
  if (taskEls.length === 0) return;

  for (var t = 0; t < taskEls.length; t++) {
    var taskEl = taskEls[t];

    // Update classes
    taskEl.classList.remove('is-done', 'is-cancelled');
    if (data.newType === 'done') taskEl.classList.add('is-done');
    if (data.newType === 'cancelled') taskEl.classList.add('is-cancelled');

    // Update checkbox icon (preserve checklist vs task distinction)
    var cb = taskEl.querySelector('.tz-task-cb');
    if (cb) {
      var isCL = data.isChecklist || cb.classList.contains('checklist');
      cb.className = 'tz-task-cb ' + data.newType + (isCL ? ' checklist' : '');
      var icon = cb.querySelector('i');
      if (icon) {
        if (isCL) {
          if (data.newType === 'done') icon.className = 'fa-solid fa-square-check';
          else if (data.newType === 'cancelled') icon.className = 'fa-solid fa-square-minus';
          else icon.className = 'fa-regular fa-square';
        } else {
          if (data.newType === 'done') icon.className = 'fa-solid fa-circle-check';
          else if (data.newType === 'cancelled') icon.className = 'fa-solid fa-circle-minus';
          else icon.className = 'fa-regular fa-circle';
        }
      }
    }
  }

  showToast(data.newType === 'done' ? 'Task completed' : data.newType === 'cancelled' ? 'Task cancelled' : 'Task reopened');
}

function handlePriorityChanged(data) {
  // Full refresh since priority changes multiple elements
  triggerRefresh();
}

function handleTaskScheduled(data) {
  // Full refresh since schedule changes grouping
  triggerRefresh();
}

function triggerRefresh() {
  sendMessageToPlugin('refresh', {
    query: currentQuery,
    filterId: currentFilterId,
    groupBy: currentGroupBy,
  });
}

// ============================================
// SEARCH / FILTER HANDLING
// ============================================

function handleSearchSubmit() {
  var input = document.querySelector('.tz-search-input');
  if (!input) return;
  currentQuery = input.value.trim();
  // If query changed from the original, we're in "edited" mode but keep the filter context
  updateSaveButtonVisibility();
  sendMessageToPlugin('runFilter', {
    query: currentQuery,
    filterId: currentFilterId,
    groupBy: currentGroupBy,
  });
}

function handleSearchClear() {
  var input = document.querySelector('.tz-search-input');
  if (input) input.value = '';
  currentQuery = 'open overdue';
  currentFilterId = '__overdue';
  originalQuery = currentQuery;
  updateSaveButtonVisibility();
  sendMessageToPlugin('runFilter', {
    query: currentQuery,
    filterId: currentFilterId,
    groupBy: currentGroupBy,
  });
}

function handleFilterClick(filterItem) {
  currentQuery = filterItem.dataset.query || 'open';
  currentFilterId = filterItem.dataset.filterId || '';
  originalQuery = currentQuery;
  closeMobileSidebar();
  updateSaveButtonVisibility();
  // Don't send groupBy — let the plugin restore the saved per-filter preference
  sendMessageToPlugin('runFilter', {
    query: currentQuery,
    filterId: currentFilterId,
    groupBy: null,
  });
}

function handleGroupByClick(btn) {
  currentGroupBy = btn.dataset.group || 'note';
  sendMessageToPlugin('runFilter', {
    query: currentQuery,
    filterId: currentFilterId,
    groupBy: currentGroupBy,
  });
}

// ============================================
// FILTER DELETE
// ============================================

function handleFilterDelete(deleteBtn) {
  var filterId = deleteBtn.dataset.deleteId;
  if (!filterId) return;
  sendMessageToPlugin('deleteFilter', {
    filterId: filterId,
    currentQuery: currentQuery,
    groupBy: currentGroupBy,
  });
}

// ============================================
// SAVE FILTER MODAL
// ============================================

function updateSaveButtonVisibility() {
  var saveArea = document.querySelector('.tz-save-area');
  if (!saveArea) return;
  var input = document.querySelector('.tz-search-input');
  var currentVal = input ? input.value.trim() : currentQuery;
  var queryChanged = currentVal !== originalQuery && currentVal !== '';
  saveArea.style.display = queryChanged ? 'flex' : 'none';
}

function isSavedFilter(filterId) {
  return filterId && filterId.startsWith('f_');
}

function showSaveDropdown() {
  closeAllPickers();
  var saveBtn = document.querySelector('.tz-save-btn');
  if (!saveBtn) return;
  var input = document.querySelector('.tz-search-input');
  var query = input ? input.value.trim() : currentQuery;
  if (!query) return;

  // If not editing a saved filter, go straight to "Save as new"
  if (!isSavedFilter(currentFilterId)) {
    showSaveAsNewModal(query);
    return;
  }

  // Show dropdown with two options
  var dropdown = document.createElement('div');
  dropdown.className = 'tz-save-dropdown';

  var updateOpt = document.createElement('button');
  updateOpt.className = 'tz-save-dropdown-opt';
  updateOpt.innerHTML = '<i class="fa-solid fa-floppy-disk"></i> Save changes';
  updateOpt.addEventListener('click', function(e) {
    e.stopPropagation();
    sendMessageToPlugin('updateFilter', {
      filterId: currentFilterId,
      query: query,
      groupBy: currentGroupBy,
    });
    originalQuery = query;
    updateSaveButtonVisibility();
    dropdown.remove();
  });
  dropdown.appendChild(updateOpt);

  var newOpt = document.createElement('button');
  newOpt.className = 'tz-save-dropdown-opt';
  newOpt.innerHTML = '<i class="fa-solid fa-plus"></i> Save as new filter';
  newOpt.addEventListener('click', function(e) {
    e.stopPropagation();
    dropdown.remove();
    showSaveAsNewModal(query);
  });
  dropdown.appendChild(newOpt);

  // Position relative to save button
  saveBtn.parentElement.style.position = 'relative';
  saveBtn.parentElement.appendChild(dropdown);

  // Close on outside click
  setTimeout(function() {
    document.addEventListener('click', function closeDropdown() {
      dropdown.remove();
      document.removeEventListener('click', closeDropdown);
    }, { once: true });
  }, 10);
}

function showSaveAsNewModal(query) {
  var overlay = document.createElement('div');
  overlay.className = 'tz-modal-overlay';

  var modal = document.createElement('div');
  modal.className = 'tz-modal';

  var title = document.createElement('div');
  title.className = 'tz-modal-title';
  title.textContent = 'Save as New Filter';
  modal.appendChild(title);

  var nameInput = document.createElement('input');
  nameInput.type = 'text';
  nameInput.className = 'tz-modal-input';
  nameInput.placeholder = 'Filter name...';
  nameInput.autofocus = true;
  modal.appendChild(nameInput);

  var queryDisplay = document.createElement('div');
  queryDisplay.className = 'tz-modal-query';
  queryDisplay.textContent = 'Query: ' + query;
  modal.appendChild(queryDisplay);

  var actions = document.createElement('div');
  actions.className = 'tz-modal-actions';

  var cancelBtn = document.createElement('button');
  cancelBtn.className = 'tz-modal-btn';
  cancelBtn.textContent = 'Cancel';
  cancelBtn.addEventListener('click', function() { overlay.remove(); });
  actions.appendChild(cancelBtn);

  var saveBtn = document.createElement('button');
  saveBtn.className = 'tz-modal-btn primary';
  saveBtn.textContent = 'Save';
  saveBtn.addEventListener('click', function() {
    var name = nameInput.value.trim();
    if (!name) {
      nameInput.style.borderColor = 'var(--tz-red)';
      return;
    }
    sendMessageToPlugin('saveFilter', {
      name: name,
      query: query,
      groupBy: currentGroupBy,
    });
    overlay.remove();
  });
  actions.appendChild(saveBtn);

  modal.appendChild(actions);
  overlay.appendChild(modal);
  document.body.appendChild(overlay);

  overlay.addEventListener('click', function(e) {
    if (e.target === overlay) overlay.remove();
  });

  nameInput.addEventListener('keydown', function(e) {
    if (e.key === 'Enter') saveBtn.click();
    if (e.key === 'Escape') overlay.remove();
  });

  setTimeout(function() { nameInput.focus(); }, 50);
}

// ============================================
// TASK ACTIONS
// ============================================

function handleTaskAction(actionEl) {
  var action = actionEl.dataset.action;
  var taskEl = actionEl.closest('.tz-task');
  if (!taskEl || !action) return;

  var encodedFilename = taskEl.dataset.encodedFilename;
  var lineIndex = parseInt(taskEl.dataset.lineIndex, 10);

  switch (action) {
    case 'toggleComplete':
      sendMessageToPlugin('toggleTaskComplete', {
        encodedFilename: encodedFilename,
        lineIndex: lineIndex,
      });
      break;

    case 'toggleCancel':
      sendMessageToPlugin('toggleTaskCancel', {
        encodedFilename: encodedFilename,
        lineIndex: lineIndex,
      });
      break;

    case 'cyclePriority':
      sendMessageToPlugin('cycleTaskPriority', {
        encodedFilename: encodedFilename,
        lineIndex: lineIndex,
      });
      break;

    case 'showSchedule':
      showSchedulePicker(actionEl, encodedFilename, lineIndex);
      break;

    case 'assignPerson':
      showAssignPicker(actionEl, encodedFilename, lineIndex);
      break;

    case 'openNote':
      sendMessageToPlugin('openNote', {
        encodedFilename: encodedFilename,
      });
      break;
  }
}

// ============================================
// SCHEDULE PICKER
// ============================================

function showSchedulePicker(anchorEl, encodedFilename, lineIndex) {
  closeAllPickers();

  var picker = document.createElement('div');
  picker.className = 'tz-sched-picker';

  var options = [
    { label: 'Today', value: todayStr() },
    { label: 'Tomorrow', value: tomorrowStr() },
    { label: 'This week', value: thisWeekStr() },
    { label: 'Next week', value: nextWeekStr() },
  ];

  for (var i = 0; i < options.length; i++) {
    var opt = document.createElement('button');
    opt.className = 'tz-sched-opt';
    opt.textContent = options[i].label;
    opt.dataset.dateValue = options[i].value;
    opt.addEventListener('click', function(e) {
      e.stopPropagation();
      sendMessageToPlugin('scheduleTask', {
        encodedFilename: encodedFilename,
        lineIndex: lineIndex,
        dateStr: this.dataset.dateValue,
      });
      closeAllPickers();
    });
    picker.appendChild(opt);
  }

  // Custom date input
  var dateInput = document.createElement('input');
  dateInput.type = 'date';
  dateInput.className = 'tz-sched-date-input';
  dateInput.addEventListener('change', function(e) {
    e.stopPropagation();
    if (!this.value) return;
    sendMessageToPlugin('scheduleTask', {
      encodedFilename: encodedFilename,
      lineIndex: lineIndex,
      dateStr: this.value,
    });
    closeAllPickers();
  });
  picker.appendChild(dateInput);

  // Clear schedule
  var clearOpt = document.createElement('button');
  clearOpt.className = 'tz-sched-opt danger';
  clearOpt.textContent = 'Remove schedule';
  clearOpt.addEventListener('click', function(e) {
    e.stopPropagation();
    sendMessageToPlugin('scheduleTask', {
      encodedFilename: encodedFilename,
      lineIndex: lineIndex,
      dateStr: '',
    });
    closeAllPickers();
  });
  picker.appendChild(clearOpt);

  // Position picker using fixed coordinates
  document.body.appendChild(picker);
  var rect = anchorEl.getBoundingClientRect();
  picker.style.top = (rect.bottom + 4) + 'px';
  picker.style.left = Math.min(rect.left, window.innerWidth - 170) + 'px';

  // Close picker on outside click
  setTimeout(function() {
    document.addEventListener('click', closePickerOnOutsideClick);
  }, 10);
}

// ============================================
// ASSIGN PICKER
// ============================================

function showAssignPicker(anchorEl, encodedFilename, lineIndex) {
  closeAllPickers();

  var picker = document.createElement('div');
  picker.className = 'tz-sched-picker tz-assign-picker';

  // Search/filter input
  var searchInput = document.createElement('input');
  searchInput.type = 'text';
  searchInput.className = 'tz-assign-search';
  searchInput.placeholder = 'Type name...';
  picker.appendChild(searchInput);

  // Options container
  var optionsDiv = document.createElement('div');
  optionsDiv.className = 'tz-assign-options';
  picker.appendChild(optionsDiv);

  var mentions = (typeof allMentions !== 'undefined') ? allMentions : [];

  function renderOptions(filterText) {
    while (optionsDiv.firstChild) optionsDiv.removeChild(optionsDiv.firstChild);
    var ft = (filterText || '').toLowerCase().replace(/^@/, '');
    var shown = 0;

    for (var i = 0; i < mentions.length; i++) {
      var m = mentions[i];
      var mClean = m.replace(/^@/, '').toLowerCase();
      if (ft && mClean.indexOf(ft) === -1) continue;

      var opt = document.createElement('button');
      opt.className = 'tz-sched-opt';
      opt.textContent = m;
      opt.dataset.mention = m;
      opt.addEventListener('click', function(e) {
        e.stopPropagation();
        doAssign(this.dataset.mention);
      });
      optionsDiv.appendChild(opt);
      shown++;
    }

    // Check for exact match
    var exactMatch = false;
    if (ft) {
      for (var j = 0; j < mentions.length; j++) {
        if (mentions[j].replace(/^@/, '').toLowerCase() === ft) { exactMatch = true; break; }
      }
    }

    // Show "create new" option if filter text doesn't exactly match an existing mention
    if (ft && !exactMatch) {
      var newOpt = document.createElement('button');
      newOpt.className = 'tz-sched-opt new-mention';
      newOpt.textContent = 'Assign to @' + ft;
      newOpt.addEventListener('click', function(e) {
        e.stopPropagation();
        doAssign('@' + ft);
      });
      optionsDiv.appendChild(newOpt);
    }
  }

  function doAssign(mention) {
    sendMessageToPlugin('assignPerson', {
      encodedFilename: encodedFilename,
      lineIndex: lineIndex,
      mention: mention,
      query: currentQuery,
      filterId: currentFilterId,
      groupBy: currentGroupBy,
    });
    closeAllPickers();
  }

  // Filter on input
  searchInput.addEventListener('input', function() {
    renderOptions(this.value.trim());
  });

  // Enter to assign custom value
  searchInput.addEventListener('keydown', function(e) {
    if (e.key === 'Enter') {
      e.preventDefault();
      var val = this.value.trim();
      if (val) doAssign(val);
    }
    if (e.key === 'Escape') closeAllPickers();
  });

  // Initial render
  renderOptions('');

  // Position and show
  document.body.appendChild(picker);
  var rect = anchorEl.getBoundingClientRect();
  picker.style.top = (rect.bottom + 4) + 'px';
  picker.style.left = Math.min(rect.left, window.innerWidth - 200) + 'px';

  // Focus the search input
  setTimeout(function() { searchInput.focus(); }, 50);

  // Close on outside click
  setTimeout(function() {
    document.addEventListener('click', closePickerOnOutsideClick);
  }, 10);
}

function closeAllPickers() {
  document.querySelectorAll('.tz-sched-picker').forEach(function(p) { p.remove(); });
  document.removeEventListener('click', closePickerOnOutsideClick);
}

function closePickerOnOutsideClick(e) {
  if (!e.target.closest('.tz-sched-picker')) {
    closeAllPickers();
  }
}

// ============================================
// FILTER DRAG-AND-DROP REORDER
// ============================================

var dragSrcEl = null;

function initFilterDragAndDrop() {
  var savedItems = document.querySelectorAll('.tz-filter-item.saved');
  savedItems.forEach(function(item) {
    item.addEventListener('dragstart', handleDragStart);
    item.addEventListener('dragover', handleDragOver);
    item.addEventListener('dragenter', handleDragEnter);
    item.addEventListener('dragleave', handleDragLeave);
    item.addEventListener('drop', handleDrop);
    item.addEventListener('dragend', handleDragEnd);
  });
}

function handleDragStart(e) {
  dragSrcEl = this;
  this.classList.add('is-dragging');
  e.dataTransfer.effectAllowed = 'move';
  e.dataTransfer.setData('text/plain', this.dataset.filterId);
}

function handleDragOver(e) {
  e.preventDefault();
  e.dataTransfer.dropEffect = 'move';
  // Show top or bottom indicator based on cursor position
  var rect = this.getBoundingClientRect();
  var midY = rect.top + rect.height / 2;
  this.classList.remove('drag-over-top', 'drag-over-bottom');
  if (e.clientY < midY) {
    this.classList.add('drag-over-top');
  } else {
    this.classList.add('drag-over-bottom');
  }
}

function handleDragEnter(e) {
  e.preventDefault();
}

function handleDragLeave(e) {
  this.classList.remove('drag-over-top', 'drag-over-bottom');
}

function handleDrop(e) {
  e.preventDefault();
  e.stopPropagation();
  if (!dragSrcEl || dragSrcEl === this) {
    this.classList.remove('drag-over-top', 'drag-over-bottom');
    return;
  }

  // Determine insert position
  var rect = this.getBoundingClientRect();
  var midY = rect.top + rect.height / 2;
  var insertBefore = e.clientY < midY;

  // Move the DOM element
  var parent = this.parentNode;
  if (insertBefore) {
    parent.insertBefore(dragSrcEl, this);
  } else {
    parent.insertBefore(dragSrcEl, this.nextSibling);
  }

  this.classList.remove('drag-over-top', 'drag-over-bottom');

  // Collect new order and send to plugin
  var newOrder = [];
  parent.querySelectorAll('.tz-filter-item.saved').forEach(function(item) {
    newOrder.push(item.dataset.filterId);
  });
  sendMessageToPlugin('reorderFilters', { orderedIds: newOrder });
}

function handleDragEnd(e) {
  // Clean up all drag classes
  document.querySelectorAll('.tz-filter-item.saved').forEach(function(item) {
    item.classList.remove('is-dragging', 'drag-over-top', 'drag-over-bottom');
  });
  dragSrcEl = null;
}

// ============================================
// MOBILE SIDEBAR TOGGLE
// ============================================

function toggleMobileSidebar() {
  var sidebar = document.querySelector('.tz-sidebar');
  var backdrop = document.querySelector('.tz-sidebar-backdrop');
  if (!sidebar) return;
  var isOpen = sidebar.classList.contains('open');
  if (isOpen) {
    closeMobileSidebar();
  } else {
    sidebar.classList.add('open');
    if (backdrop) backdrop.classList.add('open');
  }
}

function closeMobileSidebar() {
  var sidebar = document.querySelector('.tz-sidebar');
  var backdrop = document.querySelector('.tz-sidebar-backdrop');
  if (sidebar) sidebar.classList.remove('open');
  if (backdrop) backdrop.classList.remove('open');
}

// ============================================
// TOAST
// ============================================

function showToast(message) {
  document.querySelectorAll('.tz-toast').forEach(function(t) { t.remove(); });
  var toast = document.createElement('div');
  toast.className = 'tz-toast';
  toast.textContent = message;
  document.body.appendChild(toast);
  setTimeout(function() { toast.remove(); }, 3000);
}

// ============================================
// EVENT LISTENER SETUP
// ============================================

function attachAllEventListeners() {
  // Initialize state from DOM
  var searchInput = document.querySelector('.tz-search-input');
  if (searchInput) {
    currentQuery = searchInput.value || 'open';
    originalQuery = searchInput.dataset.originalQuery || currentQuery;
  }

  var activeFilter = document.querySelector('.tz-filter-item.active');
  if (activeFilter) {
    currentFilterId = activeFilter.dataset.filterId || '__all';
  }

  // Hide save button initially (shown only when query is edited)
  updateSaveButtonVisibility();

  var activeGroup = document.querySelector('.tz-group-btn.active');
  if (activeGroup) {
    currentGroupBy = activeGroup.dataset.group || 'note';
  }

  // Search input — Enter to submit, input to track changes
  if (searchInput) {
    searchInput.addEventListener('keydown', function(e) {
      if (e.key === 'Enter') {
        e.preventDefault();
        handleSearchSubmit();
      }
    });
    searchInput.addEventListener('input', function() {
      updateSaveButtonVisibility();
    });
  }

  // Search clear button
  var clearBtn = document.querySelector('.tz-search-clear');
  if (clearBtn) {
    clearBtn.addEventListener('click', function(e) {
      e.stopPropagation();
      handleSearchClear();
    });
  }

  // Filter sidebar items
  document.querySelectorAll('.tz-filter-item').forEach(function(item) {
    item.addEventListener('click', function(e) {
      // Check if delete button was clicked
      if (e.target.closest('.tz-filter-delete')) {
        e.stopPropagation();
        handleFilterDelete(e.target.closest('.tz-filter-delete'));
        return;
      }
      handleFilterClick(item);
    });
  });

  // Group-by buttons
  document.querySelectorAll('.tz-group-btn').forEach(function(btn) {
    btn.addEventListener('click', function() {
      handleGroupByClick(btn);
    });
  });

  // Save button
  var saveBtnEl = document.querySelector('.tz-save-btn');
  if (saveBtnEl) {
    saveBtnEl.addEventListener('click', function(e) {
      e.stopPropagation();
      showSaveDropdown();
    });
  }


  // Sidebar toggle (mobile)
  var sidebarToggle = document.querySelector('.tz-sidebar-toggle');
  if (sidebarToggle) {
    sidebarToggle.addEventListener('click', function(e) {
      e.stopPropagation();
      toggleMobileSidebar();
    });
  }

  // Sidebar backdrop (mobile) — close on tap
  var backdrop = document.querySelector('.tz-sidebar-backdrop');
  if (backdrop) {
    backdrop.addEventListener('click', function() {
      closeMobileSidebar();
    });
  }

  // Initialize drag-and-drop on saved filters
  initFilterDragAndDrop();

  // Task action delegation — use event delegation on the body for task-related clicks
  document.querySelector('.tz-body').addEventListener('click', function(e) {
    var actionEl = e.target.closest('[data-action]');
    if (actionEl && actionEl.closest('.tz-task')) {
      e.stopPropagation();
      handleTaskAction(actionEl);
    }
  });
}

// Initialize
document.addEventListener('DOMContentLoaded', function() {
  attachAllEventListeners();
});
