Option Explicit



'************************************************************************
' VBA Module: modBaseline (Code) - Sample data for baseline operations.
'       Copyright (c) 2018 Lynker Technologies, LLC
'       All Rights Reserved
'
' Lynker Technologies, LLC
' 5485 Conestoga Ct, Suite 220
' Boulder, CO  80301
' 303-284-8627
' http://www.lynkertech.com
'
'
'   Version History:
' Version             Date            Author              Reason
'----------------------------------------------------------------------
' V1.00           05/08/2018      rwolvington    Original
' V1.01           05/15/2018      rwolvington    Changed comments, added something.
'
'************************************************************************



'***************************************************************
'  Sub ConvertToValues() - Converts all formulas on current
' worksheet to values (no more formulas).
'
'       Called by: any.
'       Calls: nothing.
'       Parameters: nothing but acts on activesheet.
'
'  Created by: rwolvington   05/08/2018
'
'  See also
'***************************************************************
Sub ConvertToValues()

    ActiveSheet.UsedRange.Copy
    ActiveSheet.PasteSpecial Paste:=xlValues
    

End Sub 'ConvertToValues()



'***************************************************************
'  Sub SelectUsedRange() - Select the used range on the current worksheet.
'
'       Called by:
'       Calls:
'       Parameters:
'
'  Created by: rwolvington   05/15/2018
'
'  See also
'***************************************************************
Sub SelectUsedRange()

    ActiveSheet.UsedRange.Select

End Sub 'SelectUsedRange()