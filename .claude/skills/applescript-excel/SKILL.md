---
name: applescript-excel
description: Automate Microsoft Excel on macOS using AppleScript. Use when user needs to read/write/format Excel data, create formulas, manipulate worksheets, or perform any Excel automation. Focuses on Excel-native patterns - formulas over calculations, built-in features over manual logic.
---

# AppleScript Excel Automation

## When to Use This Skill

- User mentions "Excel", "spreadsheet", "workbook", "worksheet", or ".xlsx"
- Tasks involving reading/writing/formatting Excel data
- Creating formulas, charts, pivot tables
- Any Excel automation on macOS

## Core Philosophy: Work Like a Human Excel User

**CRITICAL**: Use Excel's built-in features, not programming logic.

✅ **DO**:
- Set formulas and let Excel calculate
- Use Excel's Sort, Filter, AutoFilter features
- Leverage Excel functions (SUM, VLOOKUP, IF, etc.)
- Set ranges at once (batch operations)

❌ **DON'T**:
- Loop through cells to calculate values
- Implement sorting/filtering algorithms
- Do math in AppleScript that Excel formulas can do

## Quick Start: Execute AppleScript

You can run AppleScript directly using the command line:
```bash
osascript -e 'tell application "Microsoft Excel"
    tell active sheet
        set value of cell "A1" to "Hello"
    end tell
end tell'
```

Or use the helper script:
```bash
python scripts/execute_applescript.py "your AppleScript code here"
```

## Common Patterns

### Read Cell Value
```applescript
tell application "Microsoft Excel"
    tell active sheet
        set myValue to value of cell "A1"
    end tell
end tell
```

### Write Cell Value
```applescript
tell application "Microsoft Excel"
    tell active sheet
        set value of cell "A1" to "Hello World"
    end tell
end tell
```

### Set Formula (MOST IMPORTANT)
```applescript
tell application "Microsoft Excel"
    tell active sheet
        -- Single cell
        set formula of cell "C1" to "=A1+B1"
        
        -- Entire range (Excel auto-adjusts!)
        set formula of range "C2:C100" to "=A2+B2"
        -- Result: C2 gets =A2+B2, C3 gets =A3+B3, etc.
    end tell
end tell
```

### Find Last Row
```applescript
tell application "Microsoft Excel"
    tell active sheet
        set lastRow to first row index of (get end of used range direction toward the bottom)
    end tell
end tell
```

### Read Entire Range (Fast!)
```applescript
tell application "Microsoft Excel"
    tell active sheet
        -- Returns 2D array
        set allData to value of range "A1:C10"
    end tell
end tell
```

### Write Entire Range (Fast!)
```applescript
tell application "Microsoft Excel"
    tell active sheet
        -- Must be 2D array format: {{row1}, {row2}, {row3}}
        set value of range "A1:A3" to {{"Value1"}, {"Value2"}, {"Value3"}}
    end tell
end tell
```

## Critical Performance Rules

### 1. Batch Operations - Set Ranges, Not Individual Cells

**SLOW** (1000 operations):
```applescript
repeat with i from 1 to 1000
    set value of cell ("A" & i) to i
end repeat
```

**FAST** (1 operation):
```applescript
-- Build data list first
set myData to {}
repeat with i from 1 to 1000
    set end of myData to {i}  -- Note: {i} not i (2D format)
end repeat
-- Write once
set value of range "A1:A1000" to myData
```

### 2. Use Formulas Instead of Loops

**SLOW** (calculating in AppleScript):
```applescript
repeat with i from 2 to 1000
    set a to value of cell ("A" & i)
    set b to value of cell ("B" & i)
    set value of cell ("C" & i) to (a * b)
end repeat
```

**FAST** (let Excel calculate):
```applescript
set formula of range "C2:C1000" to "=A2*B2"
```

### 3. Read Data Once, Process in Memory
```applescript
-- Read all data at once
set allData to value of range "A1:Z1000"

-- Process in AppleScript (no more Excel calls needed)
repeat with row in allData
    -- process row
end repeat
```

## Workflow Steps

1. **Understand the task**: What's the final Excel output?
2. **Check what's open**: Which workbooks/sheets are available?
3. **Plan Excel features**: Which formulas or features to use?
4. **Execute incrementally**: Test on small data first
5. **Verify results**: Check output matches expectations

## Common Tasks with Examples

### Add a Total Row
```applescript
tell application "Microsoft Excel"
    tell active sheet
        set lastRow to first row index of (get end of used range direction toward the bottom)
        set totalRow to lastRow + 1
        set formula of cell ("C" & totalRow) to "=SUM(C2:C" & lastRow & ")"
    end tell
end tell
```

### Add Calculated Column
```applescript
tell application "Microsoft Excel"
    tell active sheet
        set lastRow to first row index of (get end of used range direction toward the bottom)
        -- Header
        set value of cell "D1" to "Total"
        -- Formula for all rows
        set formula of range ("D2:D" & lastRow) to "=B2*C2"
    end tell
end tell
```

### Add Percentage Column
```applescript
tell application "Microsoft Excel"
    tell active sheet
        set lastRow to first row index of (get end of used range direction toward the bottom)
        set totalRow to lastRow + 1
        
        -- Add total
        set formula of cell ("C" & totalRow) to "=SUM(C2:C" & lastRow & ")"
        
        -- Add percentage (absolute reference to total)
        set value of cell "D1" to "Percentage"
        set formula of range ("D2:D" & lastRow) to "=C2/$C$" & totalRow
        set number format of range ("D2:D" & lastRow) to "0.00%"
    end tell
end tell
```

## Important File Rules

- **ALL data is in Excel files that are ALREADY OPEN**
- **NEVER create new workbooks or files**
- Work with existing open workbooks only
- Explore what's available first before asking user

## When You Need More Information

### For Detailed AppleScript Syntax
See [REFERENCE.md](REFERENCE.md) for:
- Complete syntax guide
- Performance optimization details
- Advanced patterns
- Troubleshooting

### For Excel Formula Patterns
See [EXCEL-FORMULAS.md](EXCEL-FORMULAS.md) for:
- Common formula patterns (SUM, VLOOKUP, IF, etc.)
- Lookup functions
- Date calculations
- Conditional logic
- Array formulas

### For Uncertain Syntax
Use the `sdef_explorer_agent` tool to query Excel's documentation:
- "How do I set a formula in a range?"
- "What's the syntax for AutoFilter?"
- "How do I get the used range?"

## Critical Reminders Checklist

- [ ] Use formulas, not calculations
- [ ] Batch operations (set ranges, not cells)
- [ ] Test on small data first
- [ ] Files are already open
- [ ] Let Excel do the work

## Examples in Action

**User**: "Add a column that calculates tax (8%) on the price column"

**Good Response**:
```applescript
tell application "Microsoft Excel"
    tell active sheet
        set lastRow to first row index of (get end of used range direction toward the bottom)
        
        -- Add header
        set value of cell "D1" to "Tax"
        
        -- Set formula (assuming price is in column C)
        set formula of range ("D2:D" & lastRow) to "=C2*0.08"
        
        -- Format as currency
        set number format of range ("D2:D" & lastRow) to "$#,##0.00"
    end tell
end tell
```

**User**: "Sort the data by column B in descending order"

**Good Response**:
```applescript
tell application "Microsoft Excel"
    tell active sheet
        set dataRange to used range
        sort dataRange key1 (column 2 of dataRange) order1 sort descending
    end tell
end tell
```

---

**Remember**: Excel is powerful. Your job is to *configure* Excel features, not *replace* them with code logic.