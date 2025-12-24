---
name: applescript-word
description: Automate Microsoft Word on macOS using AppleScript. Use when user needs to create/edit documents, format text, add paragraphs/tables/images/headers/footers, insert page breaks, find and replace text, apply styles, or perform any Word automation tasks.
---

# AppleScript Word Automation

## When to Use This Skill

- User mentions "Word", "document", ".docx", ".doc"
- Tasks involving creating/editing documents
- Adding paragraphs, tables, images, headers, footers
- Formatting text, applying styles, character/paragraph formatting
- Find and replace operations
- Any Word automation on macOS

## Core Philosophy: Work with Text Objects and Ranges

Word automation is about:
- Working with **text objects** (the document's content container)
- Creating and manipulating **text ranges** (selections of text)
- Applying formatting via **font object** and **paragraph format**
- Understanding the correct property syntax

## Quick Start: Execute AppleScript
```bash
osascript -e 'tell application "Microsoft Word"
    make new document
end tell'
```

Or use the helper:
```bash
python scripts/execute_applescript.py "your AppleScript code"
```

## Critical Syntax Rules

### create range MUST include document reference

**ALWAYS** use `create range active document start X end Y` or `create range oDoc start X end Y`
```applescript
-- CORRECT
tell application "Microsoft Word"
    set oRange to create range active document start 0 end 0
end tell

-- ALSO CORRECT (with variable)
tell application "Microsoft Word"
    set oDoc to active document
    set oRange to create range oDoc start 0 end 0
end tell
```

### Paragraph format properties use "of" syntax
```applescript
-- CORRECT: Use "set PROPERTY of paragraph format"
tell application "Microsoft Word"
    set alignment of paragraph format of selection to align paragraph center
    set space before of paragraph 1 of active document to 12
    set space after of paragraph 1 of active document to 12
end tell
```

### Font object properties use "of" syntax
```applescript
-- CORRECT: Use "set PROPERTY of font object"
tell application "Microsoft Word"
    set name of font object of selection to "Arial"
    set font size of font object of selection to 14
end tell
```

## Working with Documents

### Create New Document
```applescript
tell application "Microsoft Word"
    set newDoc to make new document
end tell
```

### Open Document
```applescript
tell application "Microsoft Word"
    open "Macintosh HD:Users:Shared:MyDocument.doc"
end tell
```

### Save Document
```applescript
-- Save existing document
tell application "Microsoft Word"
    save active document
end tell

-- Save as new file
tell application "Microsoft Word"
    save as active document file name "Sample.doc"
end tell
```

### Close Document
```applescript
tell application "Microsoft Word"
    close active document saving yes
end tell
```

## Working with Text

### Set Document Content
```applescript
tell application "Microsoft Word"
    set content of text object of active document to "New content"
end tell
```

### Insert Text at Position
```applescript
tell application "Microsoft Word"
    set oRange to create range active document start 0 end 0
    set content of oRange to "Title"
end tell
```

### Insert Text at End
```applescript
tell application "Microsoft Word"
    insert text " the end" at end of text object of active document
end tell
```

### Type Paragraph Break
```applescript
tell application "Microsoft Word"
    type paragraph selection
end tell
```

## Formatting Text

### Font Object Properties (WORKING SYNTAX)
```applescript
tell application "Microsoft Word"
    -- Set font name
    set name of font object of selection to "Arial"
    
    -- Set font size
    set font size of font object of selection to 14
    
    -- NOTE: Color, bold, italic do NOT work reliably
    -- Only name and font size work consistently
end tell
```

### Apply Font to Range
```applescript
tell application "Microsoft Word"
    set myRange to create range active document start 0 end 100
    set name of font object of myRange to "Arial"
    set font size of font object of myRange to 24
end tell
```

### Paragraph Formatting (WORKING SYNTAX)
```applescript
tell application "Microsoft Word"
    -- Set alignment
    set alignment of paragraph 1 of active document to align paragraph center
    
    -- Set spacing
    set space before of paragraph 1 of active document to 12
    set space after of paragraph 1 of active document to (inches to points inches 0.5)
    
    -- Set indents (on paragraph format object)
    set paragraph format left indent of paragraph format of selection to (inches to points inches 0.5)
end tell
```

### Alignment Constants
- `align paragraph left`
- `align paragraph center`
- `align paragraph right`
- `align paragraph justify`

### Line Spacing
```applescript
tell application "Microsoft Word"
    -- Use the space commands
    space 1 selection  -- Single spacing
    space 15 selection -- 1.5 line spacing
    space 2 selection  -- Double spacing
end tell
```

## Working with Text Ranges

### Creating Ranges (MUST include document)
```applescript
tell application "Microsoft Word"
    -- From start/end positions
    set oRange to create range active document start 0 end 0
    
    -- For paragraphs
    set myRange to create range active document ¬
        start (start of content of text object of paragraph 1 of active document) ¬
        end (end of content of text object of paragraph 3 of active document)
end tell
```

### Formatting a Range
```applescript
tell application "Microsoft Word"
    set oRange to create range active document start 0 end 0
    set content of oRange to "Title"
    
    -- Extend range to include the text
    set oRange to change end of range oRange by a word item extend type by selecting
    
    -- Format it
    set name of font object of oRange to "Arial"
    set font size of font object of oRange to 24
end tell
```

### Formatting Paragraphs via Range
```applescript
tell application "Microsoft Word"
    set myRange to create range active document ¬
        start (start of content of text object of paragraph 1 of active document) ¬
        end (end of content of text object of paragraph 3 of active document)
    
    set alignment of paragraph format of myRange to align paragraph justify
end tell
```

### Collapsing a Range
```applescript
tell application "Microsoft Word"
    set myRange to create range active document start 0 end 100
    collapse range myRange direction collapse start
end tell
```

## Working with Tables (CORRECT SYNTAX FROM DOCS)

### Create Table
```applescript
tell application "Microsoft Word"
    set oDoc to active document
    set oTable to make new table at oDoc with properties ¬
        {text object:(create range oDoc start 0 end 0), ¬
        number of rows:3, number of columns:4}
end tell
```

### Insert Text into Table Cell
```applescript
tell application "Microsoft Word"
    if (count of tables of active document) ≥ 1 then
        set content of text object of (get cell from table table 1 of active document ¬
            row 1 column 1) to "Cell 1, 1"
    end if
end tell
```

### Create and Fill Table
```applescript
tell application "Microsoft Word"
    set oDoc to active document
    set oTable to make new table at oDoc with properties ¬
        {text object:(create range oDoc start 0 end 0), ¬
        number of rows:3, number of columns:4}
    
    set iCount to 1
    repeat with oCell in (get cells of text object of oTable)
        insert text ("Cell " & iCount) at text object of oCell
        set iCount to iCount + 1
    end repeat
    
    -- Apply formatting
    auto format table oTable table format table format colorful2 ¬
        with apply borders, apply font and apply color
end tell
```

### Get Table Cell Content (Without End Marker)
```applescript
tell application "Microsoft Word"
    set oTable to table 1 of active document
    repeat with aCell in (get cells of row 1 of oTable)
        set myRange to create range active document ¬
            start (start of content of text object of aCell) ¬
            end ((end of content of text object of aCell) - 1)
        display dialog (get content of myRange)
    end repeat
end tell
```

### Get All Table Cell Contents
```applescript
tell application "Microsoft Word"
    if (count of tables of active document) ≥ 1 then
        set oTable to table 1 of active document
        set aCells to {}
        repeat with oCell in (get cells of text object of oTable)
            set myRange to text object of oCell
            set myRange to move end of range myRange by a character item count -1
            set end of aCells to content of myRange
        end repeat
    end if
end tell
```

### Convert Text to Table
```applescript
tell application "Microsoft Word"
    -- Insert tab-delimited text
    set oRange1 to create range active document start 0 end 0
    set content of oRange1 to "one" & tab & "two" & tab & "three" & tab
    set oRange1 to change end of range oRange1 by a paragraph item ¬
        extend type by selecting
    
    -- Convert to table
    set oTable1 to convert to table oRange1 separator separate by tabs ¬
        number of rows 1 number of columns 3
end tell
```

## Find and Replace (CORRECT SYNTAX)

### Finding Text and Selecting It

When find object is from **selection**, the selection changes when text is found:
```applescript
tell application "Microsoft Word"
    set selFind to find object of selection
    set forward of selFind to true
    set wrap of selFind to find stop
    set content of selFind to "Hello"
    execute find selFind
end tell
```

**Alternative with execute find arguments:**
```applescript
tell application "Microsoft Word"
    execute find find object of selection find text "Hello" wrap find find stop ¬
        with match forward
end tell
```

### Finding Text Without Changing Selection

When find object is from **text range**, selection doesn't change:
```applescript
tell application "Microsoft Word"
    set theFind to find object of text object of active document
    tell theFind
        set content to "blue"
        set forward to true
        set myFind to execute find
    end tell
    -- myFind is true if found, selection unchanged
end tell
```

**Alternative with execute find arguments:**
```applescript
tell application "Microsoft Word"
    set myRange to text object of active document
    execute find find object of myRange find text "blue" with match forward
end tell
```

### Find and Replace

Use the **replacement object** for replace operations:
```applescript
tell application "Microsoft Word"
    set selFind to find object of selection
    tell selFind
        clear formatting
        set content to "hi"
        clear formatting replacement
        set content of replacement to "hello"
        execute find wrap find find continue replace replace all with match forward
    end tell
end tell
```

### Find with Options
```applescript
tell application "Microsoft Word"
    set selFind to find object of selection
    tell selFind
        clear formatting
        set content to "search term"
        set forward to true
        set match case to true
        set match whole word to true
        set wrap to find stop
        execute find
    end tell
end tell
```

### Replace All Occurrences
```applescript
tell application "Microsoft Word"
    tell find object of text object of active document
        clear formatting
        set content to "old"
        clear formatting replacement
        set content of replacement to "new"
        execute find replace replace all
    end tell
end tell
```

## Headers and Footers (CORRECT SYNTAX)

### Add Header Using get header
```applescript
tell application "Microsoft Word"
    set s1 to section 1 of active document
    set content of text object of (get header s1 index header footer primary) ¬
        to "Header text"
end tell
```

### Add Footer Using get footer
```applescript
tell application "Microsoft Word"
    set s1 to section 1 of active document
    set content of text object of (get footer s1 index header footer primary) ¬
        to "Footer text"
end tell
```

### Header/Footer Index Constants
- `header footer primary` - Main header/footer
- `header footer first page` - First page header/footer
- `header footer even pages` - Even pages header/footer

### Add Header and Footer Together
```applescript
tell application "Microsoft Word"
    set s1 to section 1 of active document
    
    -- Set header
    set content of text object of (get header s1 index header footer primary) ¬
        to "Document Header"
    
    -- Set footer
    set content of text object of (get footer s1 index header footer primary) ¬
        to "Document Footer"
end tell
```

### Add Page Number to Footer
```applescript
tell application "Microsoft Word"
    set s1 to section 1 of active document
    set footerObj to get footer s1 index header footer primary
    
    set content of text object of footerObj to "Page "
    insert page number at end of text object of footerObj
end tell
```

### Different First Page Header
```applescript
tell application "Microsoft Word"
    set s1 to section 1 of active document
    
    -- Different first page
    set different first page header footer of s1 to true
    
    -- Set first page header
    set content of text object of (get header s1 index header footer first page) ¬
        to "First Page Header"
    
    -- Set regular header
    set content of text object of (get header s1 index header footer primary) ¬
        to "Regular Header"
end tell
```

## Document Structure

### Insert Page Break
```applescript
tell application "Microsoft Word"
    insert break at end of text object of active document break type page break
end tell
```

### Insert Section Break
```applescript
tell application "Microsoft Word"
    insert break at end of text object of active document ¬
        break type section break next page
end tell
```

## Common Workflows

### Create Formatted Document (FROM DOCS)
```applescript
tell application "Microsoft Word"
    set oRange to create range active document start 0 end 0
    set content of oRange to "Title"
    set oRange to change end of range oRange by a word item extend type by selecting
    
    -- Format title
    set name of font object of oRange to "Arial"
    set font size of font object of oRange to 24
    
    -- Add paragraph break
    type paragraph selection
    
    -- Center and add space
    set alignment of paragraph 1 of active document to align paragraph center
    set space after of paragraph 1 of active document to (inches to points inches 0.5)
end tell
```

### Create Table with Data
```applescript
tell application "Microsoft Word"
    set oDoc to active document
    
    -- Create table
    set salesTable to make new table at oDoc with properties ¬
        {text object:(create range oDoc start 0 end 0), ¬
        number of rows:4, number of columns:3}
    
    -- Fill header row
    set content of text object of (get cell from table salesTable row 1 column 1) to "Region"
    set content of text object of (get cell from table salesTable row 1 column 2) to "Sales"
    set content of text object of (get cell from table salesTable row 1 column 3) to "Growth"
    
    -- Fill data
    set content of text object of (get cell from table salesTable row 2 column 1) to "North"
    set content of text object of (get cell from table salesTable row 2 column 2) to "$1.2M"
    set content of text object of (get cell from table salesTable row 2 column 3) to "15%"
end tell
```

### Batch Find and Replace
```applescript
tell application "Microsoft Word"
    set replacements to {{"old1", "new1"}, {"old2", "new2"}}
    
    repeat with pair in replacements
        tell find object of text object of active document
            clear formatting
            set content to item 1 of pair
            clear formatting replacement
            set content of replacement to item 2 of pair
            execute find replace replace all
        end tell
    end repeat
end tell
```

### Create Document with Headers and Footers
```applescript
tell application "Microsoft Word"
    make new document
    
    -- Add content
    set oRange to create range active document start 0 end 0
    set content of oRange to "Document Title" & return & return & "Body text here."
    
    -- Add header
    set s1 to section 1 of active document
    set content of text object of (get header s1 index header footer primary) ¬
        to "Company Name - Confidential"
    
    -- Add footer with page number
    set footerObj to get footer s1 index header footer primary
    set content of text object of footerObj to "Page "
    insert page number at end of text object of footerObj
end tell
```

## Units and Measurements

### Converting Units
```applescript
-- Inches to points
set points to inches to points inches 0.5

-- Centimeters to points
set points to centimeters to points centimeters 2.5

-- Points to inches
set inches to points to inches points 36
```

### Setting Margins
```applescript
tell application "Microsoft Word"
    set iMargin to left margin of page setup of active document
    set iMargin to iMargin + (inches to points inches 0.5)
    set left margin of page setup of active document to iMargin
end tell
```

## Selecting Text

### Select Objects
```applescript
tell application "Microsoft Word"
    -- Select table
    select table 1 of active document
    
    -- Select field
    select field 1 of active document
    
    -- Select range
    set myRange to create range active document ¬
        start (start of content of text object of paragraph 1 of active document) ¬
        end (end of content of text object of paragraph 4 of active document)
    select myRange
end tell
```

### Check Selection Type
```applescript
tell application "Microsoft Word"
    if selection type of selection is selection ip then
        display dialog "Insertion point (nothing selected)"
    end if
end tell
```

## Critical Syntax Summary

### The Rules That Actually Work

1. **create range** MUST include document reference:
```applescript
   create range active document start 0 end 0
```

2. **Font properties** use "of font object":
```applescript
   set name of font object of selection to "Arial"
   set font size of font object of selection to 14
```

3. **Paragraph properties** use "of paragraph":
```applescript
   set alignment of paragraph 1 of active document to align paragraph center
```

4. **Tables** use `make new table at oDoc`:
```applescript
   make new table at oDoc with properties {text object:(create range oDoc start 0 end 0)...}
```

5. **Headers/Footers** use `get header` and `get footer`:
```applescript
   get header section 1 index header footer primary
   get footer section 1 index header footer primary
```

6. **Find object** from selection changes selection, from range doesn't:
```applescript
   find object of selection -- Changes selection
   find object of text object of active document -- Doesn't change selection
```

## What Works vs What Doesn't

### ✅ WORKS
- `set name of font object` (font name)
- `set font size of font object`
- `set alignment of paragraph`
- `set space before/after of paragraph`
- `create range active document start X end Y`
- Tables with correct syntax
- `get header`/`get footer` commands
- Find and replace with proper syntax

### ❌ DOESN'T WORK
- Font colors, bold, italic
- `create range start X end Y` (missing document)
- Direct header/footer access without get commands

## Performance Tips

1. **Store document in variable**: `set oDoc to active document`
2. **Batch operations** inside tell blocks
3. **Use ranges for bulk operations** (doesn't change UI)
4. **Clear formatting before finds**
5. **Use find on range** when selection shouldn't change

---

**Remember**: Use `get header` and `get footer` commands for headers/footers. Find object from selection changes selection; from text range doesn't. Always include document in `create range`.