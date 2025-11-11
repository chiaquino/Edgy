# documentation.py
# -*- coding: utf-8 -*-
"""
Documentation text for the Edgy Application Toolkit.
Organised by section for readability.
"""

DOC_OVERVIEW = """..."""
DOC_ELIGIBILITY = """..."""
DOC_SCORING = """..."""
DOC_RANKING = """..."""
DOC_TIPS = """..."""

# ===============================
# In-app Documentation Text
# ===============================

DOC_OVERVIEW = """
Edgy – Edge Fund Application Toolkit
====================================

What this tool does
-------------------
Edgy helps you manage the Edge Fund application process from start to finish:

1. **Eligibility**
   - Generate Word application cards (Admin View and Scorers View).
   - Generate an Excel Eligibility Marking Sheet for the eligibility team to mark.
   - Apply an editable yes/no/unsure logic to combine scorers' answers into a final eligibility decision to produce:
     * a list of eligible applications
     * a list of non-eligible applications
     * a summary sheet


2. **Scoring**
   - Create simple scorecards (no topics selected).
   - Or create scorecards by topic, so applications go to scorers who chose those topics.
   - Produce an assignment matrix, a summary report and supporting files showing how applications were distributed among scorers.

3. **Ranking**
   - Compute average score, marking variance (difference in marking among the scorers) and a final ranking based on the average score.
   - Highlight top applications visually.


Typical workflow
----------------
A typical round might look like this:

1. **Prepare eligibility**
   - Use *Eligibility → 📚 Application Cards & Marking Sheet*.
   - Generate Admin and Scorers Word files and the Eligibility Marking Sheet. Scorer files only contain selected questions.
   - Distribute the marking sheet to the eligibility team.

2. **Make the eligibility decision**
   - Collect the completed marking sheet.
   - Use *Eligibility → ⚖️ Final Decision Eligibility* to create:
     * Eligible Applications.xlsx
     * Noneligible Applications.xlsx
     * Eligibility Summary.xlsx

3. **Scoring**
   - For eligible applications only:
     - Either use *Scoring → 📝 Simple Scorecards*, or
     - *Scoring → 🎯 Scorecards by Topic* if scorers select topics.
   - Distribute scorecards to scorers.

4. **Ranking**
   - Collect all completed scoring cards.
   - Use *Ranking → 🏆 Rank Applications* to produce a single Ranking file.

You can dip into each selection separately, but this is the “full” journey from application responses to final ranking.
"""

DOC_ELIGIBILITY = """
Eligibility – Application Cards & Marking Sheet
===============================================

Menu: Eligibility → 📚 Application Cards & Marking Sheet

This page does two jobs in one go:

1. Creates Word **Application Cards** (Admin & Scorers View). 
    - Remember: for Admin View you see all the questions in the original applications. 
    - For Scorers' View, you can select which questions to select.
2. Creates the **Eligibility Marking Sheet** for the eligibility team.

Step-by-step
------------

1. **Choose a folder to save your files**
   - Click: “Choose a folder to save your files”.
   - This folder will contain:
     * Admin View (Word files)
     * Scorers View (Word files)
     * The Eligibility Marking Sheet (Excel)

2. **Import the file with the application responses**
   - Click: “Import the file with all the applications”.
   - Select the main responses spreadsheet (Excel or CSV).
   - The program reads the columns to set up the dropdown choices.

3. **Select number of applications to process**
   - Type:
     * A number (e.g. `20`) → only the first 20 applications are used.
     * Or `All` → use every application in the file.

4. **Choose key questions (columns)**
   - In the “questions” area:
     1. **Select the question showing the organisation name**  
        This is used in filenames and in the marking sheet.

     2. **Select the question showing the Application Topics (optional **  
        If you want a Topics column in the marking sheet, pick it here.  
        If your form has no topics question, tick:
        “No topics question / skip topics”.

5. **Build your Eligibility Marking Sheet**
   - **Select number of scorers per application** (1, 2 or 3).  
     This controls how many scorer columns are created in the Eligibility Marking Sheet.

   - **Options for “If Unsure, tell us why”**  
     Use the two listboxes:
       * Left side: available options
       * Right side: options that will appear in the dropdown  
     You can:
       * Move items with the → and ← buttons
       * Add custom options
       * Remove selected options from the “selected” list

   - **Options for “If No, tell us why”**  
     Same idea: choose which reasons will appear for “No” decisions.
     

6. **Select questions for the Scorer Cards**  
        You’ll see a list of all questions.  
        - These are the questions shown in the **Scorers View** Word files.  
        - The **Admin View** always contains *all* questions.  
        - You can select multiple questions, or click “Select All Questions”.

7. **Generate the cards and marking sheet**
   - Click: **“Create Application Cards and Marking Sheet”**.

What gets created?
------------------

Inside the folder you selected:

- **Admin View**  
  A Word file per application:  
  `001_GroupName.docx`, `002_GroupName.docx`, etc.  
  Contains *all* questions and answers.

- **Scorers View**  
  A Word file per application with only the questions you selected as “Scorer View” questions.

- **Eligibility Marking Sheet (Excel)**  
  One row per application, with columns like:
  - Unique Application ID (001, 002, ...)
  - Group / Organisation Name
  - Topic (if selected)
  - For each scorer:
    * Name of scorer
    * Yes / No / Unsure
    * If Unsure, tell us why (dropdown)
    * If No, tell us why (dropdown)
    * Feedback (free text)

The eligibility team fills in this sheet and sends it back to you.

Final Decision – how the logic works
====================================

Menu: Eligibility → ⚖️ Final Decision Eligibility

This page takes the completed Eligibility Marking Sheet and produces final eligibility outcomes.

1. **Import the Eligibility Marking Sheet**
   - Click “Import the Eligibility Marking Sheet”.
   - Choose the Excel file your team has filled in.

2. **Choose the location where to save eligibility results**
   - The following files will be created there:
     * Eligibility Summary.xlsx
     * Eligible Applications.xlsx
     * Noneligible Applications.xlsx

3. **(Optional) Change the marking logic**
   - Click “Change Marking Logic”.
   - This opens a small editor for the `logic_config.json` rules.
   - You can adjust how combinations of Yes / No / Unsure are turned into a final decision.

4. **Generate Final Decision**
   - Click “Generate Final Decision”.
   - For each application, the program:
     * Looks at all scorers’ answers (Yes, No, Unsure).
     * Produces:
       - Final Score (yes / no)
       - Total number of scorers
       - How many said “Unsure”.

Logic summary
-------------
- The rules are defined in `logic_config.json`.
- For 1 scorer:
  * If the only answer is “unsure”, it is treated as “yes” (soft default).
- For 2 scorers:
  * Combinations like “yes, yes”, “yes, no”, “unsure, no”, etc. are mapped according to `LOGIC_2SCORERS`.
- For 3+ scorers:
  * The program counts how many Yes / No / Unsure and compares that pattern to the rules in `LOGIC_3SCORERS`.

You can edit this file if you want to change how disagreements are resolved.
"""

DOC_SCORING = """
Scoring – Simple & By Topic
===========================

You should only use these pages *after* you have a list of Eligible Applications.

📝 Simple Scorecards
--------------------

Menu: Scoring → 📝 Simple Scorecards

Use this when you just want to distribute eligible applications fairly between scorers (no topics assignment).

Steps:
1. **Select only the Eligible Applications**
   - Usually `Eligible Applications.xlsx` created from the Final Decision Eligibility step.

2. **Choose a location where to save the Scorecards**
   - The program will save:
     * `Matrix.xlsx`
     * A `Scorers` folder with one subfolder per scorer, each containing their scorecard Excel file.
     * (Optional) a summary Word report if you tick that box.

3. **Add scorers**
   - Either:
     * Import a scorers table (Excel) and choose the column with names, then click “Load Scorers”.
     * Or type names manually and click “Add Scorer”.
   - All scorers share a single list in the listbox.

4. **How many times should each application be marked?**
   - Choose 1, 2, or 3.
   - Example: “2” means each application will appear in two different scorers’ scorecards.

5. **Scorecard columns**
   - By default:
     * “Points out of 6”
     * “Comments (min 30 words)”
   - You can add or remove columns. These become the extra columns after ID and group name in the scorecards.

6. **Generate Matrix and Scorecards**
   - Click: “Create Scorecards”.
   - The program distributes applications across scorers as fairly as possible:
     * Each application is assigned the requested number of times (if possible).
     * No scorer gets an application twice.
     * It balances the total number of application per scorer.

Outputs:
- **Matrix.xlsx**
  - Rows: scorers.
  - Columns: total applications each scorer received (and possibly other counters).
  - Useful for checking workloads.

- **Scorers / [Scorer Name] / [Scorer Name] Scorecard.xlsx**
  - Each file has:
    * Application ID
    * Organization name
    * The score columns you defined.

- **Optional summary report**
  - A Word document summarising the assignments.

🎯 Scorecards by Topic
----------------------

Menu: Scoring → 🎯 Scorecards by Topic

Use this when scorers choose topics they are comfortable scoring.

General Logic:
----------------
- You start from an **Eligible Applications** Excel file that includes:
  * An Application ID column
  * A Group/Organisation name column
  * A Topics column (comma-separated list, e.g. “housing justice, migrant rights, disability justice”).

- For each application:
  1. The program looks at all scorers and finds who matches those topics.
  2. If the application has general topics, the program:
     * Prefer scorers who match a **non-general** topic in that application.
     * Then it will search for scorers who explicitly chose a **general topic**.
     * Finally, it will distribute to low-load scorers as a fairness fallback.
  4. It assigns each application up to N scorers (N = reviews per app).
  5. It records *why* each assignment was made (topic reason).

This means:
- Apps usually go to people who chose those topics.
- General topics help share broad applications
- Workload is kept as balanced as possible.

The general topic is assigned only to scorers who either:
    - Explicitly selected the general topic (e.g., “Systemic Change"), or
    - Didn’t match any other topic but are chosen later to rebalance workloads (the fairness fallback).

So: people who chose the general topic → get general-topic applications by priority.
People who did not choose the general topic → may get some general applications only if needed to balance workloads, and those are marked as "general_balance".


What you do in the UI
---------------------

On the “Scorecards by Topic” page you:

1. **Select the Eligible Applications file (with Topics)**  
   - Must include the ID column, group name column, and topics column.

2. **Choose a location where to save Scorecards**  
   - The matrix, scorecards and reports will be saved there.
   
   **Load Topics from Eligible Application File**
   - The program scans the topics question and shows you the list of distinct topics.

   **Click general topics**
   - In the “Topics Setup” area, tick which topics count as “general”.

3. **Define scorers**
   - Add scorers in the listbox at the top (similar to Simple Scorecards).

6. **Assign Topics to Scorers**
   - Click “🧩 Create the Scorers-Topics Table”.
   - A mini spreadsheet appears where you can:
     * Mark which topics each scorer will cover.
   - You can also load a previously saved table or save the current one.

7. **How many times should each application be marked?**
   - Choose 1–3, as before.

8. **Scorecard columns**
   - Same as Simple Scorecards: add/remove columns as needed.

9. **Generate Matrix and Topic-based Scorecards**
   - The program:
     * Assigns applications using the topic logic described above.
     * Creates scorecards for each scorer.
     * Creates:
       - `Matrix by Topic.xlsx` (how many apps per scorer per topic)
       - `Applications by Topics.xlsx` (one tab per topic)
       - `Assignments by Application (Topics).xlsx` (ID, Name, Topics, Scorer 1, Scorer 2, ...). This table is useful to find where applications have been assigned.
     * Optionally creates a detailed Word report listing:
       - Applications not assigned or under-assigned
       - Any cross-topic assignments
       - Duplicates and warnings.

"""

DOC_RANKING = """
Ranking & Final Outputs
=======================

🏆 Rank Applications
--------------------

Menu: Ranking → 🏆 Rank Applications

Use this after all scorers have filled in their scorecards.

What you need:
- A **main folder** containing one subfolder per scorer.
  * Each subfolder contains `[Scorer Name] Scorecard.xlsx`.
- The original **Responses** file from the first stage.
- The questions from the Responses file you want to see in the ranking file.

Steps:
1. **Load the folder containing all scorer scorecards**
   - The app lets you choose which subfolders (scorers) to include.

2. **Load the original Responses file**
   - This is the big spreadsheet with all the original answers.
   - You then:
     * Choose which column is the organisation/group name.
     * Tick which extra columns you'd like in the final ranking (e.g. region, budget).

3. **Choose where to save the Ranking file**
   - Select a filename and location for the final Excel file.

4. **How many applications should pass this round?**
   - Type a number (e.g. 40).
   - The ranking column is computed, and the first N applications (by rank) are highlighted:
     * “Ranking” column is shaded light red.
     * “Average Score” and “Marking Variance” are shaded light purple.

The program computes: 
    
     * `Average Score` (mean of all scores per organisation)
     * `Marking Variance` (variance of the scores: ie. how different were the marks among scorers)
     * `Ranking` (1 = highest average score)

The result is a tidy ranking file that brings together:
- Original application meta-data, and
- All scoring data and basic statistics.

"""

DOC_TIPS = """
Tips & Troubleshooting
======================

General tips
------------
- Always close Excel and Word files before running steps that **save** over them.
  The app will warn you if a file is still open.
- Keep a clear folder structure, e.g.:
  - Edge Fund FRXX/
    - 01_Eligibility/
    - 02_Scoring/
    - 03_Ranking/

- When in doubt, re-run a step into a **fresh folder** so you don’t overwrite older results.

Common issues
-------------

1. **“File in Use” errors**
   - Close the Excel or Word file.
   - Try again.

2. **“Could not identify Organisation Name column”**
   - Make sure you select the correct question/column containing the group or organisation name.

3. **Empty Score or Ranking columns**
   - Check that scorers used the correct scorecard template (with a “Points out of 6” or similar score column).
   - Make sure all scorers saved their scorecards in the expected folders.

4. **Logic not giving expected eligibility results**
   - Go to: Eligibility → ⚖️ Final Decision Eligibility → “Change Marking Logic”.
   - Review the combinations and adjust if needed.

5. **Topic-based scoring looks unbalanced**
   - Verify:
     * The Topics column in your Eligible Applications file is filled correctly.
     * Scorers’ topic selections are accurate.
     * General topics are set as intended.
   - Check the:
     * `Matrix by Topic.xlsx`
     * `Assignments by Application (Topics).xlsx`
     for a detailed view.

If something looks strange, it’s often helpful to:
- Open the intermediate Excel files (Matrix, Applications by Topics, etc.).
- Check a single application and follow it through the pipeline: eligibility → scoring → ranking.
"""
