---
name: applescript-powerpoint
description: Automate Microsoft PowerPoint on macOS using AppleScript. Use when user needs to create/modify presentations, add slides, insert text/images/shapes, format content, apply transitions, run slideshows, or perform any PowerPoint automation tasks.
---

# AppleScript PowerPoint Automation

## When to Use This Skill

- User mentions "PowerPoint", "presentation", "slides", ".ppt", ".pptx"
- Tasks involving creating/editing presentations
- Adding slides, text, images, shapes, charts
- Formatting presentations, applying transitions
- Running slideshows
- Any PowerPoint automation on macOS

## Core Philosophy: Build Presentations Programmatically

PowerPoint automation is about:
- Creating slide structures with appropriate layouts
- Adding content to shapes on slides
- Applying formatting and transitions
- Building professional presentations programmatically

## Quick Start: Execute AppleScript
```bash
osascript -e 'tell application "Microsoft PowerPoint"
    make new presentation
end tell'
```

**ALWAYS compile first, then execute** (ensures syntax is correct):
```bash
# Step 1: Compile to check syntax (fails if invalid)
osacompile -o /tmp/test.scpt -e 'your AppleScript code' && \
# Step 2: Only execute if compilation succeeded
osascript -e 'your AppleScript code'
```

## Critical Presentation Concepts

### Presentations are NOT Like Documents
- **Index does NOT mean front-to-back order**
- Index refers to the order presentations were opened/created
- **NEVER assume presentation 1 is the frontmost**
- **ALWAYS use `active presentation` to reference the frontmost presentation**
```applescript
-- WRONG: Don't assume presentation 1 is frontmost
tell presentation 1
    -- might not be the one you think!
end tell

-- CORRECT: Use active presentation
tell active presentation
    -- this is always the frontmost
end tell
```

## Working with Presentations

### Create New Presentation
```applescript
tell application "Microsoft PowerPoint"
    set newPres to make new presentation
end tell
--> Returns reference to the presentation
```

### Reference Active (Frontmost) Presentation
```applescript
tell application "Microsoft PowerPoint"
    tell active presentation
        -- your code here
    end tell
end tell
```

### Open Presentation
```applescript
set thePath to choose file with prompt "Please select a presentation:"
tell application "Microsoft PowerPoint"
    open thePath
    -- Build reference to opened presentation
    set theOpenedPresentation to first presentation whose full name = (thePath as string)
end tell
```

### Save Presentation
```applescript
-- Save to original location
tell application "Microsoft PowerPoint"
    save active presentation
end tell

-- Save to new location/format
set theOutputPath to (path to desktop folder as string) & "My Preso.ppt"
tell application "Microsoft PowerPoint"
    save active presentation in theOutputPath as save as presentation
end tell
```

**Save formats**: `save as presentation`, `save as presentation template`, `save as HTML`, `save as PowerPoint show`

### Close Presentation
```applescript
tell application "Microsoft PowerPoint"
    tell active presentation
        save  -- Save first!
        close
    end tell
end tell
```

**IMPORTANT**: The `saving` parameter is ignored by PowerPoint. Always explicitly `save` before `close`.

## Working with Slides

### Create New Slide with Layout

**CRITICAL SYNTAX**: Use `layout` property with `slide layout` constants:
```applescript
tell application "Microsoft PowerPoint"
    tell active presentation
        -- Blank layout
        make new slide at end with properties {layout:slide layout blank}
        
        -- Title slide layout
        make new slide at end with properties {layout:slide layout title slide}
        
        -- Text slide layout
        make new slide at end with properties {layout:slide layout text slide}
        
        -- Section header layout
        make new slide at end with properties {layout:slide layout section header}
        
        -- Comparison layout
        make new slide at end with properties {layout:slide layout comparison}
        
        -- Content with caption layout
        make new slide at end with properties {layout:slide layout content with caption}
        
        -- Picture with caption layout
        make new slide at end with properties {layout:slide layout picture with caption}
    end tell
end tell
```

### Available Slide Layouts (TESTED & WORKING)

**Working layout constants:**
- `slide layout blank`
- `slide layout title slide`
- `slide layout text slide`
- `slide layout section header`
- `slide layout comparison`
- `slide layout content with caption`
- `slide layout picture with caption`

**NOTE**: Layouts like `slide layout title and content` and `slide layout two content` do NOT work in current PowerPoint versions.

### Count Slides
```applescript
tell application "Microsoft PowerPoint"
    set slideCount to count slides of active presentation
end tell
```

### Reference Specific Slide
```applescript
tell application "Microsoft PowerPoint"
    tell slide 2 of active presentation
        -- work with slide 2
    end tell
end tell
```

## Working with Text on Slides

### Understanding Text Structure

**CRITICAL**: Text is NOT directly in slides. Structure is:
```
Slide → Shape → Text Frame → Text Range → Content
```

### Set Text Content of Shape
```applescript
tell application "Microsoft PowerPoint"
    tell slide 2 of active presentation
        -- Set content of first shape (usually title)
        set content of text range of text frame of shape 1 to "My Slide Title"
        
        -- Set content of second shape (usually body)
        set content of text range of text frame of shape 2 to "Slide content here"
    end tell
end tell
```

### Format Text

**WORKING**: Only `underline` property works reliably for font formatting.
```applescript
tell application "Microsoft PowerPoint"
    tell slide 2 of active presentation
        -- Set text
        set content of text range of text frame of shape 2 to "Formatted Text"
        
        -- Format the text via font property
        tell font of text range of text frame of shape 2
            set font name to "Futura"
            set font size to 24
            set font color to {255, 0, 0}  -- RGB: Red
            set underline to true  -- This works
            -- NOTE: font bold, font italic, shadow do NOT work reliably
        end tell
    end tell
end tell
```

### Font Properties (What Actually Works)

**WORKING**:
- `font name` - Font family name
- `font size` - Point size
- `font color` - RGB color {R, G, B}
- `underline` - true/false

**NOT WORKING** (do not use):
- `font bold` - Does not work
- `font italic` - Does not work
- `shadow` - Does not work
```applescript
-- Example: Working font formatting
tell font of text range of text frame of shape 1
    set font name to "Arial"
    set font size to 48
    set font color to {0, 0, 255}  -- Blue
    set underline to true
end tell
```

## Adding Visual Content

### Add Picture to Slide
```applescript
set thePicturePath to (choose file with prompt "Please select a picture:") as string
tell application "Microsoft PowerPoint"
    tell slide 2 of active presentation
        -- Create picture with position
        set thePicture to make new picture at end with properties ¬
            {top:200, left position:400, lock aspect ratio:true, file name:thePicturePath}
        
        -- Scale the picture
        tell thePicture
            scale height factor 0.5 scale scale from top left with relative to original size
            scale width factor 0.5 scale scale from top left with relative to original size
        end tell
    end tell
end tell
```

### Picture Properties
- `top`: Distance from top edge (points)
- `left position`: Distance from left edge (points)
- `lock aspect ratio`: true/false (maintain proportions)
- `file name`: Path to image file (string or POSIX file)

### Scale Picture
```applescript
tell thePicture
    -- Scale to 50% of original size
    scale height factor 0.5 scale scale from top left with relative to original size
    scale width factor 0.5 scale scale from top left with relative to original size
end tell
```

### Add Shape (WORKING PATTERN)

**IMPORTANT**: Don't specify `shape type` - let it default to rectangle, then modify as needed.
```applescript
tell application "Microsoft PowerPoint"
    tell slide 1 of active presentation
        -- Create shape (defaults to rectangle)
        set rect to make new shape at end with properties ¬
            {left position:100, top:100, width:200, height:150}
        
        -- Set fill color
        set fore color of fill format of rect to {255, 0, 0}  -- Red
        
        -- Add text to shape
        set content of text range of text frame of rect to "Text in shape"
    end tell
end tell
```

### Shape Properties
- `left position`: X position from left edge
- `top`: Y position from top edge
- `width`: Shape width in points
- `height`: Shape height in points

## Slide Backgrounds

### Change Background Color
```applescript
tell application "Microsoft PowerPoint"
    tell slide 2 of active presentation
        -- Disassociate from master background
        set follow master background to false
        
        -- Set background color (RGB)
        set fore color of fill format of background to {0, 0, 255}  -- Blue
    end tell
end tell
```

### Apply Textured Background
```applescript
tell application "Microsoft PowerPoint"
    tell slide 2 of active presentation
        set follow master background to false
        
        -- Apply texture
        preset textured background texture texture blue tissue paper
    end tell
end tell
```

**Available textures**: `texture blue tissue paper`, `texture pink tissue paper`, `texture purple mesh`, `texture bouquet`, `texture cork`, `texture granite`, `texture white marble`, and more.

## Slide Show Features

### Apply Slide Transitions

**WORKING**: Basic transition effects work, but `speed` property does NOT work.
```applescript
tell application "Microsoft PowerPoint"
    tell active presentation
        set theSlideCount to count slides
        
        -- Apply transition to all slides
        repeat with a from 1 to theSlideCount
            set entry effect of slide show transition of slide a to entry effect dissolve
        end repeat
    end tell
end tell
```

### Common Transition Effects (WORKING)
- `entry effect dissolve`
- `entry effect fade`
- `entry effect wipe right`
- `entry effect wipe left`
- `entry effect push right`
- `entry effect push left`
- `entry effect cut`
- `entry effect box in`
- `entry effect box out`

**NOTE**: The `speed` property does NOT work. Do not attempt to set transition speed.

### Run Slideshow
```applescript
tell application "Microsoft PowerPoint"
    activate  -- Bring PowerPoint to front
    run slide show slide show settings of active presentation
end tell
```

### Exit Slideshow
```applescript
tell application "Microsoft PowerPoint"
    exit slide show slide show view of slide show window of active presentation
end tell
```

## Common Workflows

### Create Multi-Slide Presentation
```applescript
tell application "Microsoft PowerPoint"
    -- Create new presentation
    set newPres to make new presentation
    
    tell newPres
        -- Title slide
        set titleSlide to make new slide at end with properties ¬
            {layout:slide layout title slide}
        tell titleSlide
            set content of text range of text frame of shape 1 to "Presentation Title"
            set content of text range of text frame of shape 2 to "Subtitle or Author"
        end tell
        
        -- Content slides
        repeat with i from 1 to 5
            set contentSlide to make new slide at end with properties ¬
                {layout:slide layout text slide}
            tell contentSlide
                set content of text range of text frame of shape 1 to "Slide " & i
                set content of text range of text frame of shape 2 to "Content for slide " & i
            end tell
        end repeat
        
        -- Closing slide
        set closingSlide to make new slide at end with properties ¬
            {layout:slide layout title slide}
        tell closingSlide
            set content of text range of text frame of shape 1 to "Thank You"
        end tell
    end tell
end tell
```

### Apply Consistent Formatting to All Slides
```applescript
tell application "Microsoft PowerPoint"
    tell active presentation
        set slideCount to count slides
        
        repeat with i from 1 to slideCount
            tell slide i
                -- Format title (shape 1)
                tell font of text range of text frame of shape 1
                    set font name to "Helvetica"
                    set font size to 44
                    set font color to {0, 0, 0}
                end tell
            end tell
        end repeat
    end tell
end tell
```

### Add Agenda Slide
```applescript
tell application "Microsoft PowerPoint"
    tell active presentation
        set agendaSlide to make new slide at end with properties ¬
            {layout:slide layout text slide}
        tell agendaSlide
            set content of text range of text frame of shape 1 to "Agenda"
            set content of text range of text frame of shape 2 to ¬
                "1. Introduction" & return & ¬
                "2. Main Content" & return & ¬
                "3. Discussion" & return & ¬
                "4. Next Steps"
        end tell
    end tell
end tell
```

### Create Data-Driven Slides from List
```applescript
set dataList to {"Q1: $100K", "Q2: $150K", "Q3: $200K", "Q4: $250K"}

tell application "Microsoft PowerPoint"
    tell active presentation
        repeat with dataPoint in dataList
            set newSlide to make new slide at end with properties ¬
                {layout:slide layout text slide}
            tell newSlide
                set content of text range of text frame of shape 1 to dataPoint
                set content of text range of text frame of shape 2 to ¬
                    "Detailed information about " & dataPoint
            end tell
        end repeat
    end tell
end tell
```

## Position and Size Reference

Positions and sizes are in **points** (72 points = 1 inch):

- **Standard 16:9 slide**: 960 × 540 points
- **Standard 4:3 slide**: 720 × 540 points
- `left position`: Distance from left edge
- `top`: Distance from top edge
- `width`: Width of object
- `height`: Height of object
```applescript
-- Example: Center a 400×300 shape on 16:9 slide (960×540)
-- Left: (960-400)/2 = 280
-- Top: (540-300)/2 = 120
set centered to make new shape with properties ¬
    {left position:280, top:120, width:400, height:300}
```

## Color Reference

Colors use **RGB values (0-255)**:
```applescript
-- Primary colors
{255, 0, 0}    -- Red
{0, 255, 0}    -- Green
{0, 0, 255}    -- Blue

-- Common colors
{0, 0, 0}      -- Black
{255, 255, 255} -- White
{128, 128, 128} -- Gray

-- Custom colors
{255, 165, 0}  -- Orange
{128, 0, 128}  -- Purple
{255, 192, 203} -- Pink
```

## Performance Tips

1. **Create presentations/slides once**, then populate
2. **Minimize tell blocks** - group operations together
3. **Build content in memory**, then add to slides
4. **Use repeat loops** for bulk slide creation
5. **Set properties during creation** with `with properties {}`

## Important File Rules

- Presentations can be **created new** OR work with existing
- Files may be **already open** - use `active presentation`
- Save explicitly before closing
- Use full paths for images and file operations

## Common Patterns Summary

### Creating Content
```applescript
-- 1. Create presentation/slide
make new presentation
make new slide at end with properties {layout:slide layout text slide}

-- 2. Add text to shapes
set content of text range of text frame of shape 1 to "Title"

-- 3. Format text (only what works!)
tell font of text range of text frame of shape 1
    set font name to "Arial"
    set font size to 36
    set font color to {0, 0, 255}
    set underline to true
end tell

-- 4. Add visuals
make new picture at end with properties {file name:imagePath, top:200, left position:300}

-- 5. Add shapes (don't specify type)
set myShape to make new shape at end with properties ¬
    {left position:100, top:100, width:200, height:150}
set fore color of fill format of myShape to {255, 0, 0}
```

## When You Need More Information

### For Detailed AppleScript Syntax
See [REFERENCE.md](REFERENCE.md) for:
- Complete AppleScript syntax guide
- Performance optimization
- Error handling
- Advanced patterns

### For Uncertain PowerPoint Commands
Use the `sdef_explorer_agent` tool:
- "What slide layouts are available?"
- "How do I add animations?"
- "What transition effects exist?"

## Critical Reminders Checklist

- [ ] Use `active presentation` not `presentation 1`
- [ ] Slides are 1-indexed (slide 1 is first)
- [ ] Text path: shape → text frame → text range → content
- [ ] Only these layouts work: blank, title slide, text slide, section header, comparison, content with caption, picture with caption
- [ ] Don't specify shape type - let it default
- [ ] Only underline works for font styling (not bold/italic/shadow)
- [ ] Transition speed property doesn't work
- [ ] Positions are in points (72 = 1 inch)
- [ ] Save before close (saving parameter is ignored)
- [ ] Images need `lock aspect ratio:true` and scaling

## Quick Reference: What Works vs What Doesn't

### ✅ WORKS
- Font name, size, color, underline
- Shape creation without type specification
- Basic transition effects
- All tested slide layouts (7 total)

### ❌ DOESN'T WORK
- Font bold, italic, shadow
- Shape type specification
- Transition speed
- Some layout types (title and content, two content)

---

**Remember**: PowerPoint AppleScript support is limited. Stick to tested, working patterns. When in doubt, test with one slide before scaling up.