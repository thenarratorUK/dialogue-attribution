## README: scripter.streamlit.app

### Manuscript Markup For Narration

#### Overview

This web app is designed to mark up a manuscript for narration, as if done so manually with a highlighter (physically or electronically)   

It works only with  **Word documents (DOCX)** using international style (double) speech marks. If your manuscript is in PDF or uses British (single) speech marks (DOCX or PDF, the app at [https://quotescript.streamlit.app](https://quotescript.streamlit.app) can be used to convert the document from PDF to DOCX (correcting the quote style if necessary) or can be used to convert a British-styled DOCX to an American styled DOCX  

The app allows you to:

- Attribute each line of dialogue to the correct speaker.
- Undo a single mistake
- Assign colours to each speaker for easy reading.
- Generate a clean, formatted HTML preview of the full script.
- Save and re-load your progress between sessions.

WARNING: Streamlit puts apps to sleep if not used for a while. If you expect a break of 12+ hours away from using the app, assume that no-one else will be using the app and download the progress files manually (see Step 4)

This tool is hosted at [**https://scripter.streamlit.app**](https://scripter.streamlit.app) and runs entirely in your browser. No installation or coding knowledge is required.

---

### How to Use

#### 1. Step 0: User Verification

#### Input a ‘User Key'

- Type whatever you would like into this box. It will be prepended to most of the files you download at the end of the process (excluding the speaker colors file) and is required for the load saved progress function to work (only available up to four hours from last use)
- Do not type something overly simple - if someone else types exactly the same thing, they will be able to access what you are working on. 

---

#### 2a. Step 1: Uploading Files (First Time)

#### Upload Your Files

- The first time you use the app with a specific document, you must go through a one-time process to extract all the dialogue from the document (the ‘quotes.txt’ file)
- Click **Browse files** or drag your files into the upload boxes.
- The **DOCX file** is required.
  - Example: `BookTitle.docx`
- Once the docx has been uploaded, click ’Start Processing'
- After a short time (seconds to minutes, depending on the length of the manuscript) you will be given the following options
  - **Download Extracted Quotes TXT**: Click this to save a copy of the quotes.txt file if you intend to return to the process later and not continue immediately
  - **Restart:** Click if you are generating quotes.txt files from multiple books in sequence
  - **Continue:** Click to move to Step 2 (number 3, below)

#### 2b. Step 1: Uploading Files (Subsequent Times)

- Click **Browse files** or drag your files into the upload boxes.

  - The **DOCX file** is required.
    - Example: `BookTitle.docx`

- You will also upload:

  - **Quotes file:** `userkey-BookTitle-quotes.txt`
    - Contains dialogue lines in the format `123: Speaker: "Line of text"`
    - Helps the app recognise which lines belong to which speaker.
  - **Speaker colours file (optional):** `BookTitle-speaker_colours.json`
    - Contains a mapping of speaker names to colour codes, e.g.\
      `{ "Kal": “light grey", "Borgrim": “red" }`

---

#### 3. Step 2 — Process Unknown Speakers

- Any dialogue lines without a recognised speaker are labelled **“Unknown.”** in the quotes.txt file
- The app guides you line-by-line to correct them.
- For each “Unknown” line, you can:
  - Type the correct speaker name (e.g., “Elowen”)
  - Type or click **skip** to leave it unchanged
  - Type or click **undo** to revert the last change that wasn’t a skip - this works for a single step only
  - Type or click **exit** to progress to step 3
- A panel shows the most frequently used speaker names (these appear once a speaker has been used 10+ times) as quick-access buttons.
- Your progress is continuously stored in the session’s working JSON cache (`userkey-BookTitle.json`).

---

#### 4. Step 3 — Assign or Adjust Speaker Colours

- Each speaker name is shown alongside a colour box.
- Only the speakers without a colour selected are displayed
- As each speaker has a colour associated with it, it will disappear from the page (colours may be preallocated from an uploaded speaker-colors.json file in Step 1). 
- There are three special colours
  - None: Does not highlight in step 4, but displays all unattributed dialogue in a dark red. This makes it easier to download the final HTML, print it to PDF, and continue to manually highlight in appropriate software.\*\* This is the default colour for all speakers until changed\*\*
  - Error: Does not highlight, and reverts text back to black. When the dialogue in Step 2 is incorrectly detected (e.g. it’s a series of italicised words that are\*\* not\*\* internal dialogue) attributing the dialogue to ‘Error’ and then subsequently assigning the colour ‘Error’ to that speaker will mean that the text displays entirely normally in Step 4
  - Do Not Read: If *Do Not Read* is typed in Step 2, this colour will be automatically attributed, but it can also be selected manually. This will highlight the given text in black and change the text to a dark grey so that it’s still just about readable, but difficult to read. This is used as a reminder to not narrate a particular section if agreed with an author/publisher. 
- Please note: there are two instances of duplicate colours
  - Bright Blue and Dark Blue are the same colour. Light Blue and Navy Blue are the two alternative blues
  - Light Pink and Pale Pink are the same colour
- You have two additional options on this page:
  - If you want to edit previously allocated colours, click ‘edit speaker colours’. This will display *all* the names along with their colours, and they will not disappear as they are changed
  - If you want to move to step 4 (you *can* leave speakers without a colour), click ‘Continue'

---

#### 5. Step 4 — Generate and Review HTML Preview

- Step 4 may take some time to load. Once loaded, it will show an HTML file and a number of downloaded boxes. 
- The HTML shows:
  - All speaker dialogue with assigned colours used to highlight.
  - All formatting preserved from the DOCX.
- You can scroll through the preview in your browser.
- There are three special sections at the top of the HTML
  - Character Summary: Displays how many lines each character has, in order of appearance. This can be useful for planning and for the creation of a voice reference
  - Speaker Ranking: This displays the same information but in the order of how many lines each character has. This is to allow you to prioritise developing individual voices for the most frequent characters. Characters with a single line of dialogue are **excluded**. 
  - First Substantial Lines: This displays the first line(s) for each character, in order of appearance, for those narrators that like to use the first line in a voice reference. This **includes** characters with a single line of dialogue. 
- Below the html preview, five (or six, in rare circumstances) buttons will appear:
  - &#x20;**Download HTML file** lets you download the HTML, either for reading from directly or for converting to PDF, as per individual workflow
  - **Download Updated Speaker Colors JSON** lets you download the speaker colors file, for reupload in step 1 if continuing after the app has gone to sleep 
  - **Download Updated Quotes TXT** lets you download the updated quotes.txt file, with the “Unknown:” at the start of each line replaced with the attribution given in Step 2. 
  - In rare circumstances, an error above the HTML may appear saying that not all quotes could be matched. If this has occured, then a button for **Download Unmatched Quotes** will appear, providing index references for the lines in quotes.txt that could not be matched
  - **Return to Step 2** lets you return to Step 2 to continue attributing dialogue (typically used when moving to step 4 has been done momentarily in order to download the json and txt files for safety)
    - ***NOTE***: *If you have input ’skip’ on a particular line, returning to Step 2 will not return to the skipped lines. To return to skipped lines, download the files and reinput them in step 1.* 
  - **Clear Cache for This User** will erase everything you’ve been working on. Clicking this prevents someone else guessing your userkey and using Load Saved Progress to see what you were working on. It is recommended to always click this when you are certain you will not be returning to the app within 12 hours

---

### Loading Saved Progress

- In Steps 1-4, clicking ‘Load Saved Progress’ at the top of the page will return you to where you previously were
- In practice, this is used as follows:
  - Input your userkey in Step 0
  - When Step 1 appears, click ‘load saved progress’ (do not upload anything as this will wipe your saved progress)
- **Note that this will only work if the app has not gone to ‘sleep’, which happens after 12 hours of inactivity from any user**

---

### Practical Notes

- Use consistent speaker names throughout (e.g., “Elowen” not “Elowen the Librarian”).
- All names are normalised. Timmy, timmy, tiMmY and TIMmy will all be stored as Timmy. When multiple words are used, each word is capitalised, e.g. James The Paramedic
- In the *extremely* rare situation of the app saying no context could be found for the preview in Step 2, click or type exit, progress through to step 4, and download the speaker colors json and quotes txt files. Refresh the page, reinput your userkey, and reupload those two files along with the docx. This will recreate the preview window correctly.
- When printing the HTML (including to PDF) please ensure the settings allow for printing backgrounds, otherwise the highlights will not appear in the print. 
