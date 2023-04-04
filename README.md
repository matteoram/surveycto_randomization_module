# SurveyCTO Randomization Module

## A turnkey script to implement in a randomization module in a SurveyCTO questionnaire and address question order effects

This Python script creates a randomization module to be used in a SurveyCTO questionnaire. It follows the approach outlined in Ramina (2022).

Simply put, this program:
* Creates a csv file containing all possible permutations given a integer indicating the number of texts to be randomized;
* Creates an xlsx file replicating the 'survey' and 'choices' tabs of a standard SurveyCTO questionnaire and populates them with the required code to randomize the order of the texts.

Importantly, this script differs from the Stata code shown in the above-mentioned technical brief in two ways:
* The wording 'statements' is replaced by 'texts', given that the module can be used to randomize not only statements, but any type of string;
* The script produces not only a csv file as in the case of the Stata code, but also automates the creation of the respective Excel file to be added to the SurveyCTO questionnare. To execute this, it requires the user to enter the texts to be randomized and the SurveyCTO field type associated to the texts (e.g. 'select_one', 'select_multiple', integer, etc.).

In order to run this program seamlessly, simply insert the correct directory of the output files. The names of the output files can also be changed.

For more information regarding the limitations of this solution, please refer to the technical brief linked above.

If you find a bug or you would like to suggest an improvement, please raise an issue through the issues tab.

## Reference

Ramina, M. (2022). Tackling Question Order Effects to Improve the Accuracy of Your Survey. Technical Brief. Laterite. Available at: www.laterite.com/blog/technical-brief-tackling-question-order-effects.