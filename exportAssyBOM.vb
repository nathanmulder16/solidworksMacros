'---------------------------------------------------------------------------
' Preconditions:
' 1. Open a drawing of an assembly with a BOM.
' 2. Click the move-table icon in the upper-left corner
'    of the BOM table to open the table's PropertyManager page.
'
' Postconditions: Saves the selected BOM to c:\temp\BOMTable.xls.
'--------------------------------------------------------------------------
Dim swApp As SldWorks.SldWorks
Dim swModDoc As SldWorks.IModelDoc2
Dim swTable As SldWorks.ITableAnnotation
Dim status As Integer
 

Option Explicit

Sub Main()
    Set swApp = Application.SldWorks
    Set swModDoc = swApp.ActiveDoc
    Dim swSM As ISelectionMgr
    Set swSM = swModDoc.SelectionManager
    Set swTable = swSM.GetSelectedObject6(1, -1)
    swModDoc.ClearSelection2 (True)

    Dim swSpecTable As IBomTableAnnotation
    Set swSpecTable = swTable

    ' Save the selected BOM table to Microsoft Excel, including hidden cells and images
    status = swSpecTable.SaveAsExcel("C:\Users\nathanm\Desktop\Nathan\Testing\bomMacro\BOMTable.xls", True, True)

End Sub