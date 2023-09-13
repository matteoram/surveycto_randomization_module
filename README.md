# SurveyCTO Randomization Module

## A turnkey script addressing question order effects via a randomization module within a SurveyCTO questionnaire

This Python script creates a randomization module intended for use in a SurveyCTO questionnaire, following the approach outlined in Ramina (2022).

Simply put, this program:
* Creates a csv file containing all possible permutations based on a integer indicating the number of texts to be randomized;
* Creates an xlsx file replicating the 'survey' and 'choices' tabs found in a standard SurveyCTO questionnaire, and populates them with the required code to randomize the order of the texts.

Importantly, this script differs from the Stata code provided in the above-mentioned technical brief in two ways:
* The term 'statements' is replaced with 'texts' to emphasize that the module can randomize any type of string, not only statements;
* In addition to producing a csv file as seen in the Stata code, this script automates the creation of the corresponding Excel file that can be seamlessly integrated into the SurveyCTO questionnaire. To execute this, the user is prompted to input the texts for randomization and specify the SurveyCTO field type associated with these texts (e.g., 'select_one,' 'select_multiple,' 'integer,' etc.).

In order to run this program smoothly, simply provide the correct directory path for the output files. You also have the flexibility to customize the names of the output files.

For more information regarding the limitations of this solution, please refer to the technical brief linked above.

If you find a bug or you would like to suggest improvements, please raise an issue through the issues tab.

## Reference

Ramina, M. (2022). Tackling Question Order Effects to Improve the Accuracy of Your Survey. Technical Brief. Laterite. Available at: www.laterite.com/blog/technical-brief-tackling-question-order-effects.