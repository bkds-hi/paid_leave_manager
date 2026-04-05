# AGENTS.md - Development Guidelines

## Project Overview

Google Apps Script project for a paid leave (有給) management tool that runs in Google Sheets. Automatically calculates remaining days, manages grant dates using FIFO consumption, and tracks expiration.

## Build/Lint/Test Commands

**No build system** - This is a pure Google Apps Script project.

- **Deploy**: Copy `appsscript.js` to Google Apps Script editor
- **Run**: Execute functions directly in GAS editor or via spreadsheet triggers
- **No automated tests**: Manual testing via spreadsheet interaction
- **No linter**: Follow style guidelines below

## Code Style Guidelines

### Language & Syntax
- **ES5 JavaScript** (Google Apps Script runtime)
- Use `var` for variable declarations (not `let`/`const`)
- Use traditional `function` declarations (not arrow functions)
- 2-space indentation
- Unix line endings (LF)

### Naming Conventions
- **Functions**: `camelCase` (e.g., `autoProcess`, `updateGrantSummary`)
- **Variables**: `camelCase` (e.g., `runningTotal`, `expireDate`)
- **Constants**: `SCREAMING_SNAKE_CASE` in a `CONSTANTS` object
- **Event handlers**: Use GAS convention (e.g., `onEdit`, `onOpen`)

### Comments
- Use Japanese for user-facing documentation
- Use English for technical comments explaining complex logic
- JSDoc style for function documentation:
```javascript
/**
 * Brief description of what the function does
 * Additional details in Japanese for user-facing docs
 * 
 * @param {Type} paramName - Description
 * @returns {Type} Description
 */
function functionName(paramName) { }
```

### File Organization
Organize code in clear sections with comment separators:
```javascript
// ============================================================================
// Section Name
// ============================================================================
```

Order of sections:
1. Constants
2. Event Handlers
3. Main Processing Functions
4. Data Reading Functions
5. Calculation Functions
6. Data Writing Functions

### Constants
All magic numbers and configuration should be in `CONSTANTS` object:
```javascript
var CONSTANTS = {
  HEADER_ROW: 1,
  COL_DATE: 1,
  EXPIRY_YEARS: 2,
  STATUS_VALID: '有効'
};
```

### Error Handling
- Wrap event handlers in try-catch blocks
- Log errors using `console.error()`
- Show user-friendly error messages via `SpreadsheetApp.getUi().alert()`
- Use early returns for guard clauses
- Validate inputs before processing

### Formatting
- Opening braces on same line: `if (condition) {`
- Always use braces for control structures
- Semicolons at end of statements
- No trailing whitespace
- Maximum line length: 100 characters
- Separate logical sections with blank lines

### GAS APIs
- Use `SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()` for current sheet
- Prefer `getValues()`/`setValues()` for batch operations
- Clear ranges before writing new data to avoid stale values
- Use constants for column numbers (1-based indexing)

### Spreadsheet Conventions
- Column A-D: User input (date, delta, running total, memo)
- Column E: Separator (blank)
- Column F-K: Auto-generated summary data
- Row 1: Headers (never modify programmatically)
- Use 1-based indexing for rows/columns

## Architecture

```
onOpen()              # Adds custom menu on spreadsheet open

onEdit(e)             # Event handler for cell edits
    ↓
autoProcess()         # Main orchestration function
    ↓
├── readTransactions()        # Read data from columns A-D
├── calculateRunningTotals()  # Calculate running totals
├── writeRunningTotals()      # Write to column C
├── processGrants()           # FIFO consumption logic
├── buildGrantSummaries()     # Build summary data
└── writeGrantSummaries()     # Write to columns F-K
```

## Dependencies

No external dependencies. Uses only Google Apps Script built-in APIs:
- `SpreadsheetApp`
- `SpreadsheetApp.getUi()`

## Testing

Manual testing workflow:
1. Open spreadsheet with script attached
2. Input date in column A, delta in column B
3. Verify column C updates automatically
4. Check F-K columns for grant summary
5. Use "有給管理" menu → "再計算" for manual recalculation
6. Test edge cases: negative numbers, future dates, expired grants

## Deployment

1. Open [Google Apps Script](https://script.google.com)
2. Create new project or open existing spreadsheet
3. Paste `appsscript.js` contents
4. Save and authorize permissions
5. Run `onOpen()` once to add menu
6. Test in linked spreadsheet

## File Structure

```
holiday/
├── appsscript.js      # Main GAS code (entry point)
├── README.md          # Project overview and usage guide
├── DESIGN.md          # Architecture documentation
├── AGENTS.md          # Development guidelines (this file)
├── LICENSE            # MIT License
└── .gitignore         # Git ignore patterns
```

## Contributing

1. Follow ES5 syntax requirements
2. Maintain the existing code organization structure
3. Add JSDoc comments for new functions
4. Test changes manually in a Google Sheet
5. Update documentation if behavior changes
