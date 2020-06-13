# Random Tables 

## What is this?  

Random Tables came out of an exploration of solo tabletop RPG. With the absense of a GM, the soloist takes on activities of both player and GM.  

To aid in the GM side of things, a soloist can use random tables. The soloist roles a die, like a d20, and uses the result to reference a table whose entries are numbered, in this case 1 to 20.  

![old-school random table](/images/keep_table.jpg)

Perhaps this table has entries of magicial items, rolled when opening a chest. Or perhaps the table has entries of monster encounters, rolled when proceeding through a wilderness.  

The soloist often records the experience as a narrative in a document or journal. The narrative can be as descriptive or as prosy as the soloist desires, enough to express the events of the adventure.  

A popular solo tabletop RPG is [Ironsworn](https://www.ironswornrpg.com/), by Shawn Tomkin. Many players use Google Docs to journal their adventures.

Google Sheets is a natural fit for storing and working with tables.  

I saw an opportunity to combine the adventure journal in Google Docs with the automation the rolling of random tables in Google Sheets. 

The middleware connecting the two is Google App Script.

## What does this do?

After this Add-on is installed, **Random Tables** appears under the Add-ons menu. Select the **Show sidebar** option under the **Random Tables** Add-on menu to show the sidebar.  In the sidebar, use the load control to load a Google Sheet from a url. The Google Sheet must follow the format below. After the load completes, the sidebar will be populated with buttons that call functions (random tables) in the Google Sheet, and write the result into the document.

Examples of buttons are: 
1. Roll d20 - return a number between 1 and 20 
2. Roll d6 - return a number between 1 and 6
3. Roll Treasure - return a random row from the Treasure table
4. Oracle Roll - return the answer to a Yes/No question from the Oracle table

Google Sheets is a very powerful tool, and not every function needs to employ a table. For example, the *Roll d20* button can use the RANDBETWEEN() Google Sheet function to generate a random number.  

These functions may take input from the user. If an input is defined for the function, the button will prompt the user with a dialog.  Google Sheet data validation for range and for list are supported and are rendered as a dropdown lists in the dialog.

![](/images/screen1.png)

## Google Sheet Format

A Google Sheet (spreadsheet) can contain one or more named sheets. The following sections describe a sheet named **Index** and a sheet named **Links**.

For Google Sheet examples of the format described below, see the example folder URL in the section [**How do I use this?**](#how-do-i-use-this)

### Index Sheet
The sidebar buttons are generated from entries in a sheet named **Index**.  The first row is assumed to be a header and is skipped.  

Columns A to C contain the following data:  
* A - (required) the button text
* B - (required) a string representing the (single) output cell, in A1 notation
* C - (optional) a string representing the range of input cells, in A1 notation

The input range is a range of two rows. The first row is the input label, the second row is the input value. Each column of the range represents a single input argument.  

#### Index Sheet Example
|      | A        | B            | C               |
| :--- | :---     | :---         | :---            |
| 1    | Button   | Output       | Input           |
| 2    | Roll d20 | Functions!A1 | Functions!A2:A3 |

In the above example, the button defined at row 2 will show the text *Roll d20*. The output cell contains a string that references the *A1* cell in the sheet called *Functions*. The input cell contains a string that references a range from A2 to A3 in the sheet called *Functions*.  

#### Functions Sheet Example
|      | A                  | B     |
| :--- | :---               | :---  |
| 1    | ="d20+"&A3&": "&A5+A3  |       |
| 2    | Modifier           |       |
| 3    | 1                  |       |
| 4    | Roll               |       |
| 5    | =RANDBETWEEN(1,20) |       |

In the above example, the output cell is A1, where the formula builds an output string. The input takes a single parameter in the range A2:A3, where A2 is the input label and A3 is the input value. If there were an additional input parameters, they would start in the next column at B2:B3. 

Upon successful load from the sidebar, a section of buttons is added to the sidebar with the section titled from the Google Sheet filename.  

The section title supports the following GUI operations:
* The *equals button* collapses the button area
* The *x button* removes the section
* The title bar can be used to drag the section for rearranging the order of the sections

Upon successful load, the url in the sidebar textbox is saved as a document property. Currently, this is the only sidebar value saved. When the sidebar is reopened, the textbox is repopulated and the url is reloaded automatically. Loading an empty textbox will delete the saved url value from the document property.  

You may have several commonly used Google Sheets that you want to load as a group. This may be accomplished with a sheet named **Links**.

### Links Sheet

When a Google Sheet url is loaded from the sidebar, a sheet named **Links** can be used to load one or more additional Google Sheet urls.  The first row is assumed to be a header and is skipped.  

#### Links Sheet Example
|      | A        | B            | 
| :--- | :---     | :---         | 
| 1    | Name     | URL          | 
| 2    | Dice     | https://docs.google.com/spreadsheets/d/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA/edit | 
| 3    | Mythic GME     | https://docs.google.com/spreadsheets/d/BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB/edit | 

Columns A to B contain the following data:  
* A - (optional) a descriptive name
* B - (required) a string url of a Google Sheet

If any of the additional Google Sheet urls have their own **Links** sheets, these are not loaded. Only the **Links** sheet of Google Sheet url loaded directly from the sidebar will be loaded.  

## How do I use this?

This is now available as a Google Doc Add-on in the G Suite Marketplace.

#### G Suite Marketplace (option 1)
https://gsuite.google.com/marketplace/app/appname/645649794340

For examples of Google Sheets that implement the format above, see the **Random Tables Examples** folder URL below.

**Random Tables Examples**   
https://drive.google.com/drive/folders/1S30zpuHb1cNiS0uHXMJoq-lLDELhLQrB?usp=sharing

To use these files, select them all and use the **Make a copy** option to make copies to your Google Drive. Then use your copies of the files.

The above public Google Drive folder contains the following example files.

**Character - Steve (Google Sheet)**  
* A simple character sheet implemented as a spreadsheet with functions that perform specific common checks.

**Dice (Google Sheet)**  
* Several common dice rolls (some prompt for a multiplier), and an additional general dice roll where the number of sides can be specified.

**Rolls (Google Sheet)**  
* A few generic dice roll checks from known games.

**Tables (Google Sheet)**  
* A few general table rolls with different lookup techniques.

#### Copy the Source (option 2)

In a new Google Doc, use the **Tools -> Script editor** menu option. This will result in a new script project (referred to as a "container-bound script" in the Google App Script documentation.)  

You will need to recreate the files under the /src folder from this repository in the Google App Script editor. Don't create a src folder in the Google App Script editor, just put the files in the root. You won't need the src/appsscript.json file.

When you run the script for the first time in a document (by using the Add-On menu), you will be prompted to give permissions. Once you have completed the prompts, you will need to use the Add-On menu again.  This only happens once per document.  
