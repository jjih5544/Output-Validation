# Output-Validation
This Python script is a desktop application built using Tkinter and the Smartsheet API. Its purpose is to allow users to input (or scan) a "Lot Number", validate whether that number already exists in a specified Smartsheet column, and if it doesn't, add it to the top of the sheet.

Function: GUI app for validating and adding unique Values to a Smartsheet.

**Technology stack:**
<br>Tkinter: for GUI
<br>Smartsheet: for interacting with Smartsheet API
<br>Threading: for keeping the UI responsive during background operations
<br>
<br>Key Features:
1. Prevents duplicate entries
2. Displays loading spinner during background work
3. Shows popup messages for feedback (success, errors, etc.)

**How It Works (Step-by-Step)**
<br>1. On Launch:
<br>Loads existing Lot Numbers into cached_values.
<br>Starts the worker thread to listen for new submissions.
<br>2. User Enters Value:
<br>Submits by pressing Enter or clicking "Submit".
<br>Value is pushed into submission_queue.
<br>3.Worker Processes Submission:
<br>Checks if the Lot Number already exists.
<br>If new:
<br>Shows loading spinner.
<br>Adds to Smartsheet.
<br>Adds to cached_values.
<br>Shows a success message.
<br>If duplicate:
<br>Shows info popup.
<br>4.User Sees Feedback:
<br>Success / duplicate / error popups guide the user.
