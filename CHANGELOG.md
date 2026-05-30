# What's changed in 🔍 Task Zoom plugin?

## [1.1.1] 2026-04-12
### New
- **Wiki-link rendering**: `[[Note Name]]` references in task text become clickable links that open the note.
- **Split-view note opening**: clicking a note title opens it in a split view alongside the dashboard.
- **Repeating-task support**: completing a repeating task invokes the Routine plugin to handle recurrence; a setting controls this behaviour.
- **`@done` timestamp**: completing a task from the UI appends `@done(date+time)` to the task.
- **Saved-filter context menu**: right-click a saved filter for rename, duplicate, and delete actions; new filter button added to the toolbar.
- **Parenthesis support in query parser**: complex boolean expressions such as `(#work | #home) & !done` now parse correctly.
- **Inline and end-line comment styling**: comments inside task text are visually distinguished.

### Changes
- Priority colors are now read dynamically from the active NotePlan theme (`flagged-1/2/3`) so they stay in sync with light and dark themes.
- Filter and group-by switching is instant via a pre-built HTML cache; a loading spinner covers slower first loads.
- All duplicate instances of a task across notes are updated together when completing or cancelling.

### Fixes
- Checklist items no longer convert to regular tasks when completed or cancelled; their checkbox style is preserved.
- Checklist items are correctly included/excluded by filters.
- Per-filter group-by preference is properly restored when switching between saved filters.
- Mobile layout: priority badge and task title stay on the same line.
- Save-filter button remains visible after editing a query with the Enter key.

## [1.0.0] 2026-03-21
- Initial release: **Task Zoom** command — a smart task-filter dashboard with Todoist-like query syntax, saved filters, flexible grouping (folder, note, status, tag, mention, date, priority), light/dark theme support, assign picker, and checklist-item display.
