# README
## Module 2 - VBA Challenge
### Process Stock Data

The following sites were used to complete this VBA project. 

- https://www.automateexcel.com/vba/loops/
- https://www.perplexity.ai/search/Given-this-code-pgKJptjlRxCbbil9mjCXlQ
- https://stackoverflow.com/questions/44409090/percent-style-formatting-in-excel-vba
- https://stackoverflow.com/questions/13661965/conditional-formatting-using-excel-vba-code.

### Discussion of approach and citations
I originally approached this challange using the GetUniqueValues function discussed here: https://stackoverflow.com/questions/31690814/how-do-i-get-a-list-of-unique-values-from-a-range-in-excel-vba. However after working through some of the exercises  on Monday and Tuesday, and having difficulties parsing date column between the long and short file sets, I realized I didn't need the date column for the actual Multiple_year_stock_data file because the tabs were broken down by quarter. On Tuesday, 25JUN2024, the substitute instructor provided the link to https://www.automateexcel.com/vba/loops/, which I used along with the perplexity AI discussion about the lotto numbers exercise (https://www.perplexity.ai/search/Given-this-code-pgKJptjlRxCbbil9mjCXlQ), to formulate my approach to the loop for the VBA Challenge. I used the stackoverflow discussion at https://stackoverflow.com/questions/44409090/percent-style-formatting-in-excel-vba, to format the K column values. In the section formatting positive and negative values, use of the FormatConditions commands to delete previous formatting, highlight positive values green, and negative value red, was taken from https://stackoverflow.com/questions/13661965/conditional-formatting-using-excel-vba-code. In addition, intializing the following variables: outputRow = 2, tickerStartRow = 2, totalVolume = 0, outside of the inner loop for each sheet came from a suggestion by our substitute TA Brandon when working on the checkerboard exercise and the census 1 exercise. The idea for a summaryRow (or outputRow in my script) also came from TA Brandon Wong while working on an in group exercise on Thursday evening, 27JUN2024. 
