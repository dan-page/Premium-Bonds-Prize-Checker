# Premium-Bonds-Prize-Checker
A Python-based automated prize checker for NS&amp;I Premium bonds which automatically logs winnings to an excel file.

It does this by checking the past 6 months of winning bonds, found here: https://www.nsandi.com/get-to-know-us/winning-bonds-downloads.

To Run the code:
1. Install the required packages by running: pip install -r requirements.txt
2. Save the Winnings.xlsx Excel file to the required location
3. Update the NS&I Holdings sheet with the bond numbers that want to be checked
4. Open premiumBondsChecker.py and update the path and filename for the Excel file
5. The code is now ready to run. It should automatically identify which months winnings are missing and then update the excel accordingly.


