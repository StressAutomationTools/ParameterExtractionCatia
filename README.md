# ParameterExtractionCatia
Catia Macro to extract Parameters to a text file

Requires Microsoft Scripting Runtime as reference
Writest parameters to a text file in C:\Temp\CatParameters.txt (overwriting existing file).
The text file is tab separated for easy input in excel: ParameterName\tParameterValue

Note: The script will work on the first opened part or product only. If the open file is neither a part nor a product, it will give an error message.
