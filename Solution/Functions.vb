
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.DataSourcesGDB

Module Functions
    Public Function FieldsExist(ByVal pClass As IClass, ByVal createFieldName As String, ByVal modifyFieldName As String, ByVal userFieldName As String) As Boolean

        If (pClass.FindField(createFieldName) < 0 Or pClass.FindField(modifyFieldName) < 0 Or pClass.FindField(userFieldName)) Then
            FieldsExist = False
        Else
            FieldsExist = True
        End If

    End Function

    Public Sub CreateFields(ByVal pClass As IClass, ByVal createFieldName As String, ByVal modifyFieldName As String, ByVal userFieldName As String)

        Dim lCreated As Long
        Dim lModifiedLast As Long
        Dim lModifiedBy As Long
        Dim lPSRCguid As Long
        lCreated = pClass.FindField(createFieldName)
        lModifiedLast = pClass.FindField(modifyFieldName)
        lModifiedBy = pClass.FindField(userFieldName)
        'lPSRCguid = pClass.FindField(guidFieldName)

        If lCreated < 0 Then
            'if "Created" field does not exist, create it
            Dim pField As IField
            pField = New FieldClass
            Dim pFieldedit As IFieldEdit
            pFieldedit = pField     'QI

            'Set properties for new field
            With pFieldedit
                .Editable_2 = True
                .IsNullable_2 = True
                .Length_2 = 50
                .Name_2 = createFieldName
                .Precision_2 = 25
                .Scale_2 = 0
                .Type_2 = esriFieldType.esriFieldTypeDate
            End With
            'Add the new field
            pClass.AddField(pField)
        End If

        If lModifiedLast < 0 Then
            'If "ModifiedLast" field does not exist, create it
            Dim pField As IField
            pField = New FieldClass
            Dim pFieldEdit As IFieldEdit
            pFieldEdit = pField
            With pFieldEdit
                .Editable_2 = True
                .IsNullable_2 = True
                .Name_2 = modifyFieldName
                .Length_2 = 25
                .Scale_2 = 0
                .Type_2 = esriFieldType.esriFieldTypeDate
            End With

            pClass.AddField(pField)
        End If

        If lModifiedBy < 0 Then
            'If "ModifiedBy" field does not exist, create it
            Dim pField As IField
            pField = New FieldClass
            Dim pFieldEdit As IFieldEdit
            pFieldEdit = pField
            With pFieldEdit
                .Editable_2 = True
                .IsNullable_2 = True
                .Name_2 = userFieldName
                .Length_2 = 50
                .Type_2 = esriFieldType.esriFieldTypeString
            End With

            pClass.AddField(pField)

        End If



       






    End Sub

    Public Sub GetLargestID(ByVal pClass As IClass, ByVal AttField As String, ByRef LargestAtt As Long)

        'on error GoTo eh
        Dim pTable As ITable
        pTable = pClass
        Dim pSort As ITableSort
        Dim pCs As ICursor, pRow As IRow
        pSort = New TableSort

        'hyu - this block seems redundant, so is comment out.
        With pSort
            .QueryFilter = Nothing
            .Fields = pTable.OIDFieldName
            .Ascending(pTable.OIDFieldName) = False
            .Table = pTable
            .Sort(Nothing)
        End With

        pCs = pSort.Rows
        pRow = pCs.NextRow

        With pSort
            .QueryFilter = Nothing
            .Fields = AttField
            .Ascending(AttField) = False
            .Table = pTable

            .Sort(Nothing)
        End With

        pCs = pSort.Rows
        pRow = pCs.NextRow
        LargestAtt = pRow.Value(pTable.FindField(AttField))

        Exit Sub


eh:
        MsgBox("Error: GetLargestID " & vbNewLine & Err.Description)

    End Sub

    Public Function GenerateGuid(ByVal pWorkspace As IWorkspace) As String

        Dim pGenerateGuid As IGUIDGenerator
        pGenerateGuid = pWorkspace
        GenerateGuid = pGenerateGuid.CreateGUID

        


    End Function
    Public Function GetWorkspace(ByVal path As String, ByVal hwnd As Integer) As IWorkspace
        GetWorkspace = Nothing
        Dim pWSFact As IWorkspaceFactory
        pWSFact = New FileGDBWorkspaceFactory
        GetWorkspace = pWSFact.OpenFromFile(path, hwnd)



    End Function

End Module
