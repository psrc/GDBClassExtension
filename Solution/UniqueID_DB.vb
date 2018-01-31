Imports System.Windows.Forms
Imports System.Data.SqlClient

Public Class UniqueID_DB
    Public myConnection As SqlConnection
    Public Property EdgeID() As Long
        Get
            Return GetID("PSRCEdgeID")

        End Get
        Set(ByVal value As Long)
            SetID("PSRCEdgeID", value)



        End Set
    End Property
    Public Property OSMEdgeID() As Long
        Get
            Return GetID("OSM_PSRCEdgeID")

        End Get
        Set(ByVal value As Long)
            SetID("OSM_PSRCEdgeID", value)



        End Set
    End Property
    Public Property JunctionID() As Long
        Get
            Return GetID("PSRCJunctionID")

        End Get
        Set(ByVal value As Long)
            SetID("PSRCJunctionID", value)


        End Set
    End Property
    Public Property OSMJunctionID() As Long
        Get
            Return GetID("OSM_PSRCJunctionID")

        End Get
        Set(ByVal value As Long)
            SetID("OSM_PSRCJunctionID", value)


        End Set
    End Property
    Public Property ProjectRteID() As Long
        Get
            Return GetID("ProjectRteID")
        End Get
        Set(ByVal value As Long)
            SetID("ProjectRteID", value)


        End Set
    End Property
    Public Property OSMProjectRteID() As Long
        Get
            Return GetID("OSM_ProjectRteID")
        End Get
        Set(ByVal value As Long)
            SetID("OSM_ProjectRteID", value)


        End Set
    End Property
    Public Property TurnID() As Long
        Get
            Return GetID("TurnID")

        End Get
        Set(ByVal value As Long)
            SetID("TurnID", value)


        End Set
    End Property
    Public Function GetID(ByVal FieldName As String) As Long
        

        Dim myDataAdapter As New SqlDataAdapter
        Dim myDataSet As DataSet = New DataSet
        myConnection = New SqlConnection(My.Settings.DBConnect2)
        myDataAdapter.SelectCommand = New SqlCommand

        myDataAdapter.SelectCommand.Connection = myConnection
        ' myDataAdapter.SelectCommand.CommandText = "SET TRANSACTION ISOLATION LEVEL REPEATABLE READ GO BEGIN TRANSACTION GO SELECT * FROM UNIQUEIDS (UPDLOCK) WHERE UNIQUEIDNAME = @UniqueIDName GO COMMIT TRANSACTION"
        myDataAdapter.SelectCommand.CommandText = "SET TRANSACTION ISOLATION LEVEL REPEATABLE READ BEGIN TRANSACTION SELECT UniqueIDName, UniqueIDValue FROM dbo.tblPrimaryKeyCounters WHERE UNIQUEIDNAME = @UniqueIDName COMMIT TRANSACTION"
        'myDataAdapter.SelectCommand.CommandText = "SELECT UniqueIDName, UniqueIDValue FROM dbo.tblPrimaryKeyCounters WHERE UNIQUEIDNAME = @UniqueIDName"
        myDataAdapter.SelectCommand.Parameters.AddWithValue("@UniqueIDName", FieldName)
        myDataAdapter.Fill(myDataSet, "UniqueIDs_Table")
        Dim myDataView As DataView = New DataView(myDataSet.Tables("UniqueIDs_Table"))

        GetID = myDataView.Item(0).Item("UniqueIDValue")

        'myConnection.Close()


    End Function

    Public Sub SetID(ByVal FieldName As String, ByVal UniqueID As Long)

        'Dim myConnection As SqlConnection = New SqlConnection("Data Source=.\SQLEXPRESS;AttachDbFilename=C:\Program Files\Microsoft SQL Server\MSSQL.1\MSSQL\Data\TAZTEST.mdf;Integrated Security=True;Connect Timeout=30;User Instance=True")
        Dim myCommand As SqlCommand = New SqlCommand
        myConnection = New SqlConnection(My.Settings.DBConnect2)
        myCommand.Connection = myConnection

        myCommand.CommandText = "UPDATE dbo.tblPrimaryKeyCounters SET UniqueIDValue = @UniqueIDValue WHERE UNIQUEIDNAME = @UniqueIDName"
        myCommand.CommandType = CommandType.Text
        myCommand.Parameters.AddWithValue("@UniqueIDName", FieldName)

        myCommand.Parameters.AddWithValue("@UniqueIDValue", UniqueID)
        myCommand.Connection.Open()
        myCommand.ExecuteNonQuery()


        myCommand.Connection.Close()
        'MessageBox.Show(myConnection.State.ToString)
















    End Sub

End Class
