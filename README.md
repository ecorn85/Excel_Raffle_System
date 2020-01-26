# Excel_Raffle_System

## Purpose
With streamers taking donations for suggestions/nicknames, the purpose of this project is to give them a way to:
  1. Keep an organized list of entries.
  2. Use a randomized raffle selection system.
  3. Keep track of what entries have already been used.
  4. The ability to expand up/create new worksheets while still maintaining functionality.
  
## Implementation
The workbook uses Excel's Visual Basic API to manipulate cells. It also uses Excel's built-in formulas and conditional formatting.

The code can be accessed by pressing "Alt + F11."

## Requirements
- Excel 2007 or later
- Macros must be enabled (usually by a Security Warning propmt from Excel) for the code to run.

## How it works

### Overview
Each row acts as its own lottery pool. The count column keeps track of the number of filled cells within the pool. Clicking on the selector will look at the pool, look for unused entries, and select one at random. The selected entry is displayed to the right of the selection cell, and the entry is highlighted in yellow to indicate it has been used.

For easy cleanup, double-clicking the big red button will clear all entries out of all pools.

<br />

**Feel free to add/remove rows. The code is designed to scale with the number of rows.**

**Do not add/remove columns unless you are prepared to edit the code/formatting.**

<br />


### Detailed Explanation

<br />
#### Type Columns (if applicable)
Uses conditional formating rules to edit the cells' styles based upon the text value of the cells. This can be edited through the conditional formatting options.
#### Count Column
Keeps track of the number of filled cells in the pool for that row. (=COUNTA(RANGE))
#### Selection Column
Clicking on the cell will call upon the code to evaluate the current state of the names pool, and return a random, unused entry if one exists.
#### Selected Column
When selcting a random entry, the result will be displayed here.

<br />

Possible values:
- *Empty* = There have been no attempts to select an entry.
- *Entry* = An entry has been successfully selected, and tagged as used.
- "No Entry Available" = There are no entries in the pool.
- "All Entries Used" = All listed entries have been marked as used.
#### Entry Pool
This area stores the entries, with a different pool for each row. It is one entry per cell in the row. The code will accomodate for skipping cells, though it is not recommended.

Yellow highlight indicates that the entry has been selected, and will not be chosen again. To reenter an entry, simply remove the highlight.

#### Clear Button
Double Clicking this cell will remove all entries from all pools, and clear the *Selected* column.

**Calling code cannot be undone with an Undo.**
<br />

# Credits
Created by: ecorn85 aka Qwellfar

Inspired by: Yogscast Stream Team
