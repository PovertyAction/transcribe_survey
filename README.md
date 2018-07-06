transcribe_survey
==============
transcribe_survey is a Python program that takes ODK programming and produces a paper survey. [Zach Groff](zachary.groff@aya.yale.edu) wrote it in the spring of 2016 while a Research Analyst on the Dean Team at IPA. Zach edited the program in the spring of 2018 while a Senior Research Analyst for Dean at the Buffett Institute at Northwestern.

What the Program Does
---------------------------------
In short, transcribe_survey takes your ODK file to program a SurveyCTO survey and translates it into an English-language word document. Each row in the survey becomes a question, and the survey lists out the questions in order with headings for each group.

transcribe_survey has a few nice features:

-The program includes all aspects of a question: the question itself (or the "label" column in ODK), the format of the answer (text, number, or multiple choice), and the hint, relevance, or constraint if relevant. The formatting displays the question as it might be displayed if written by a human.
-Proper formatting for headings and instructions helps to match the way a human might design a survey.
-The program translates ODK expressions into English sentences. transcribe_survey turns mathematical expressions like <= to "is less than or equal to" and replaces references to earlier fields with blanks and parentheticals instructing surveyors to fill in the answer to the previous question. The translations also account for whether an expression is a constraint, relevance, calculation, or note so that the syntax is correct for the case.
-The program prints repeat groups multiple times as is needed. If the ODK file specifies how many times a repeat group should repeat (the "repeat_count" column in ODK), then the program repeats the group that many times. The user can also specify a default number of repeats in the case that the repeat_count column is blank. Each repeat group starts with an overall heading before a subheading for each round of the repetition.
-The program can generally print the questions from any language included in the ODK survey. Note that the translation feature does not work for other languages, so mathematical expressions will be translated into English regardless the language specified. This problem can be avoided by choosing not to display constraints, relevance, and calculations.
-The program can print any innermost repeat group as a table. That is, if a repeat group does not contain an additional repeat group, the program can print a table with each question in the repeat group as a column and each repeat as a row. This can be useful for family rosters or summaries of the quantities and prices of crops or assets.

Running the Program
-----------------------------
This program runs in Python 2.7. Before running this program, install openpyxl and python-docx. You can install these programs using [pip install](https://pip.pypa.io/en/stable/installing/).

To run this program, either download transcribe_survey.py or clone this GitHub repository to your desktop. Set the Python directory and the names of the ODK file and word document. Choose whatever options you desire, and then hit F5 or choose "Run Module" in the "Run" menu.

Program Options
------------------------
transcribe_survey offers the user options to specify how the survey appears and how comprehensive it is. The options are listed below, with the ones that must be filled out in bold. The one that may need to be filled out depending on the survey are in italics:

**excelname**?This option specifies the path and name of the ODK file to be translated.

**wordname**?This option specifies the path and name of the word document to be created from the ODK file.

*language*?This option specifies which label column in the ODK file should be used if the ODK file includes language-specific label columns. If the ODK file contains multiple languages, this option must be specified.

formtitle?This option specifies the title of the survey. If not specified, transcribe_survey will automatically use the title from the settings tab of the ODK file as the title.

suppress?This option tells the program whether to suppress repeat groups or not. If suppress=1, the program will not repeat repeat groups and will instead act as if the repeat groups are non-repeat groups. If suppress=0, the program will repeat repeat groups.

tablesinclude?This option tells the program tells the survey whether to include tables or not. If tablesinclude=1, the program will format innermost repeat groups as tables. If tablesinclude=0, the program will format all repeat group questions the same as other questions.

relevances?This option tells the program whether to include the relevance column or not. If relevances=1, the program will print relevances, translated into English, as part of hints in italics following questions. If relevances=0, the program will not print relevances.

constraints?This option tells the program whether to include the constraint column or not and behaves similarly to relevances.

calculates?This option tells the program whether to include calculate and calculate_here questions in the survey. If calculates=1, calculate and calculate_here fields will appear as blanks with instructions on the calculation for the surveyor to perform. For instance, if num_fruits and num_vegetables correspond to questions 14 and 15, respectively, then the calculation "${num_fruits}+${num_vegetables}" would become "The answer to Q14 plus the answer to Q15." A blank line would follow for the surveyor to perform the calculation. If calculates=0, calculate and calculate_here fields will be omitted.