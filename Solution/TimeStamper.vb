

Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Utility.CATIDs
Imports ESRI.ArcGIS.Framework
Imports System.IO
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Windows.Forms
Imports ESRI.ArcGIS.ADF.CATIDs





<ComClass(TimeStamper.ClassId, TimeStamper.InterfaceId, TimeStamper.EventsId), ProgId("EadaGDBExt.TimeStamper")> _
Public Class TimeStamper
    Implements IClassExtension
    Implements IObjectClassEvents
    Implements IRelatedObjectClassEvents

    Private m_class As IClass
    Private m_DataSet As IFeatureClass
    Private m_propertySet As IPropertySet

    Private m_createFieldName As String
    Private m_modifyFieldName As String
    Private m_userFieldName As String
    Private m_guidFieldName As String
    Private m_PSRCEdgeID As String

    Private m_userName As String

    Private m_UserControl As UserControl1
    Private PSRCEdgeID As Long
    Private ProjRteID As Long
    Private TurnID As Long
    Private PSRCEdgeID2 As Long
    Private EdgeCreated As Boolean = False
    Private oldPSRCEdgeID As Long
    Private time1 As TimeSpan







#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "E23D47A1-51CA-480F-B777-F581C3B3F845"
    Public Const InterfaceId As String = "A158A1FC-8507-483D-BCCC-A52B8D4330F4"
    Public Const EventsId As String = "3E459B0B-4E16-4C57-8988-B4DC31C9D656"
#End Region

#Region "Component Category Registration"
    <ComRegisterFunction()> _
    Public Shared Sub Reg(ByVal regKey As String)
        GeoObjectClassExtensions.Register(regKey)
    End Sub

    <ComUnregisterFunction()> _
    Public Shared Sub Unreg(ByVal regKey As String)
        GeoObjectClassExtensions.Unregister(regKey)
    End Sub
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub Init(ByVal pClassHelper As ESRI.ArcGIS.Geodatabase.IClassHelper, ByVal pExtensionProperties As ESRI.ArcGIS.esriSystem.IPropertySet) Implements ESRI.ArcGIS.Geodatabase.IClassExtension.Init
        ' Hold on to back reference to class and the propertyset
        m_class = pClassHelper.Class
        m_propertySet = pExtensionProperties


        If (m_propertySet.Count = 0) Then
            Exit Sub ' Extension must be initialized!
        End If

        ' Get the current user's name
        m_userName = Environ("USERNAME")

        ' Get the field names from the property set
        m_createFieldName = m_propertySet.GetProperty("CREATION_FIELDNAME")
        m_modifyFieldName = m_propertySet.GetProperty("MODIFICATION_FIELDNAME")
        m_userFieldName = m_propertySet.GetProperty("USER_FIELDNAME")
        

        ' Create the fields in feature class if they are missing
        If (Not Functions.FieldsExist(m_class, m_createFieldName, m_modifyFieldName, m_userFieldName)) Then
            Functions.CreateFields(m_class, m_createFieldName, m_modifyFieldName, m_userFieldName)
        End If

        'Optional
        m_UserControl = New UserControl1

    End Sub

    Public Sub Shutdown() Implements ESRI.ArcGIS.Geodatabase.IClassExtension.Shutdown

        m_propertySet = Nothing
        m_class = Nothing
        m_UserControl = Nothing 'Optional

    End Sub

  

    Public Sub OnChange(ByVal obj As ESRI.ArcGIS.Geodatabase.IObject) Implements ESRI.ArcGIS.Geodatabase.IObjectClassEvents.OnChange

        If (m_propertySet.Count = 0) Then Exit Sub

        ' Set the Modification date and user name
        Dim row As IRow
        row = obj

        Dim i As Integer
        i = m_class.FindField(m_modifyFieldName)
        row.Value(i) = Now

        i = m_class.FindField(m_userFieldName)
        row.Value(i) = m_userName




        ' Optional
        'Inspect(row)
    End Sub

    Public Sub OnCreate(ByVal obj As ESRI.ArcGIS.Geodatabase.IObject) Implements ESRI.ArcGIS.Geodatabase.IObjectClassEvents.OnCreate
        'Dim pApp As ESRI.ArcGIS.Framework.IApplication
        ' pApp = New AppRef

        Dim m_uniqueIDs As New UniqueIDs

        If (m_propertySet.Count = 0) Then Exit Sub



        ' Set the creation date and user name
        Dim row As IRow
        row = obj

        Dim i As Integer
        Dim x As Long

        'updates the date created field
        i = m_class.FindField(m_createFieldName)
        row.Value(i) = Now

        'updates the user field name (editor)
        i = m_class.FindField(m_userFieldName)
        row.Value(i) = m_userName

        'if the class is TransRefEdges, updtates PSRCEdgeID

        If obj.Class.AliasName = "sde.SDE.TransRefEdges" Then
            EdgeCreated = True
            i = m_class.FindField("PSRCEdgeID")
            If i <> -1 Then
                Dim newUniqueValue As UniqueID_DB = New UniqueID_DB

                row.Value(i) = newUniqueValue.EdgeID + 1

                PSRCEdgeID = row.Value(i)

                newUniqueValue.EdgeID = PSRCEdgeID

            End If
        ElseIf obj.Class.AliasName = "OSMtest.DBO.TransRefEdges" Then
            i = m_class.FindField("PSRCEdgeID")
            If i <> -1 Then

                Dim newUniqueValue As UniqueID_DB = New UniqueID_DB

                row.Value(i) = newUniqueValue.OSMEdgeID + 1

                newUniqueValue.OSMEdgeID = row.Value(i)

            End If
            'if the class is TransRefJunctions, updtates PSRCEdgeID
        ElseIf obj.Class.AliasName = "sde.SDE.TransRefJunctions" Then
            i = m_class.FindField("PSRCJunctID")
            If i <> -1 Then

                Dim newUniqueValue As UniqueID_DB = New UniqueID_DB

                row.Value(i) = newUniqueValue.JunctionID + 1

                newUniqueValue.JunctionID = row.Value(i)

            End If

        ElseIf obj.Class.AliasName = "OSMtest.DBO.TransRefJunctions" Then
            i = m_class.FindField("PSRCJunctID")
            If i <> -1 Then

                Dim newUniqueValue As UniqueID_DB = New UniqueID_DB

                row.Value(i) = newUniqueValue.OSMJunctionID + 1

                newUniqueValue.OSMJunctionID = row.Value(i)

            End If

        ElseIf obj.Class.AliasName = "sde.SDE.ProjectRoutes" Then
            i = m_class.FindField("PROJRTEID")
            If i <> -1 Then

                Dim newUniqueValue As UniqueID_DB = New UniqueID_DB

                row.Value(i) = newUniqueValue.ProjectRteID + 1

                newUniqueValue.ProjectRteID = row.Value(i)

            End If

        ElseIf obj.Class.AliasName = "OSMtest.DBO.ProjectRoutes" Then
            i = m_class.FindField("PROJRTEID")
            If i <> -1 Then

                Dim newUniqueValue As UniqueID_DB = New UniqueID_DB

                row.Value(i) = newUniqueValue.OSMProjectRteID + 1

                newUniqueValue.OSMProjectRteID = row.Value(i)

            End If


        ElseIf obj.Class.AliasName = "sde.SDE.TurnMovements" Then
            i = m_class.FindField("TurnID")
            If i <> -1 Then

                Dim newUniqueValue As UniqueID_DB = New UniqueID_DB

                row.Value(i) = newUniqueValue.TurnID + 1

                newUniqueValue.TurnID = row.Value(i)

            End If

        End If




    End Sub

    Public Sub OnDelete(ByVal obj As ESRI.ArcGIS.Geodatabase.IObject) Implements ESRI.ArcGIS.Geodatabase.IObjectClassEvents.OnDelete

    End Sub







    ' Optional
    Private Sub Inspect(ByVal row As IRow)

        

    End Sub

    Public Sub RelatedObjectCreated(ByVal RelationshipClass As ESRI.ArcGIS.Geodatabase.IRelationshipClass, ByVal objectThatWasCreated As ESRI.ArcGIS.Geodatabase.IObject) Implements ESRI.ArcGIS.Geodatabase.IRelatedObjectClassEvents.RelatedObjectCreated

        Dim pFeatureDataset As IFeatureDataset
        Dim pTable1 As ITable
        Dim pTable2 As ITable
        Dim x As Integer
        Dim pOID1 As Integer
        Dim pOID2 As Integer

        Try

            pFeatureDataset = RelationshipClass.FeatureDataset

            If RelationshipClass.OriginClass.AliasName = "sde.SDE.TransRefEdges" Or RelationshipClass.OriginClass.AliasName = "OSMtest.DBO.TransRefEdges" Then

                'get the PSRCEdgeID from the feature that was created
                pOID1 = objectThatWasCreated.OID

                pTable1 = RelationshipClass.OriginClass

                x = pTable1.FindField("PSRCEDGEID")

                PSRCEdgeID = pTable1.GetRow(pOID1).Value(x)

                Dim pcursor As ICursor
                Dim prow As IRow



                'create a new record and assign it the ProjectRteID
                pTable2 = RelationshipClass.DestinationClass

                x = pTable2.FindField("PSRCEDGEID")

                pOID2 = pTable2.CreateRow.OID

                pTable2.GetRow(pOID2).Value(x) = PSRCEdgeID

                pTable2.GetRow(pOID2).Store()

            End If

            If RelationshipClass.OriginClass.AliasName = "sde.SDE.ProjectRoutes" Or RelationshipClass.OriginClass.AliasName = "OSMtest.DBO.ProjectRoutes" Then

                'get the ProjectRteID from the feature ethat was created
                pOID1 = objectThatWasCreated.OID

                pTable1 = RelationshipClass.OriginClass

                x = pTable1.FindField("PROJRTEID")

                ProjRteID = CType(pTable1.GetRow(pOID1).Value(x), Long)

                'MessageBox.Show(test)

                'create a new record and assign it the ProjectRteID
                pTable2 = RelationshipClass.DestinationClass

                x = pTable2.FindField("PROJRTEID")

                pOID2 = pTable2.CreateRow.OID

                pTable2.GetRow(pOID2).Value(x) = ProjRteID

                'MessageBox.Show(pTable2.GetRow(pOID2).Value(x))

                pTable2.GetRow(pOID2).Store()

            End If


        Catch ex As Exception

            MessageBox.Show(ex.ToString & vbCrLf & vbCrLf & "Please contact Stefan Coe if problem persists.", "Data Catalog", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        End Try


    End Sub


    

    
End Class