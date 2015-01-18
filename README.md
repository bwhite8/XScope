# XScope
MS Excel Add-In with GUI for Automated Reporting

## Please Note
This tool was built and designed for a specific system & database. Sensitive content, especially that involving reference to SQL tables has been removed for security reasons. Removed content will be denoted by "%REMOVED%"

## User Install:
 - Save XScope.xlam in Excel's root add-in folder.  Often (Windows 7 Default) this directory is:
   C:\Users\yourname\Documents\AppData\Roaming\Microsoft\AddIns

 - Open Microsoft Excel.  You should notice a new tab in your primary ribbon labeled "XScope".  If not, open your Add-In
   Manger and ensure the checkbox labeled "XScope" is checked.

## For Your Convenience
VBA code has been exported from the add-in and loaded separately for your viewing pleasure.  
 - frmXScope.frm : GUI
 - modXScope.bas : Handles userform controls, builds & executes SQL queries, operates on ADODB recordsets, returns data
   to GUI



