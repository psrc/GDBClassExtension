Attribute VB_Name = "AlterClassExtCLSID"
Option Explicit

Public Sub AlterGDBExtCLSID()

'Takes the selected feature class in ArcCatalog and changes the EXTCLSID value of this feature class to that of the
'class extension's component's GUID value.

Dim gxApp As IGxApplication
Dim gxObject As IGxObject

Set gxApp = Application
Set gxObject = gxApp.SelectedObject

'make sure something is selected
If (Not (gxObject Is Nothing)) Then

'make sure data is selected
    If (Not TypeOf gxObject Is IGxDataset) Then
        MsgBox "You must first specify the feature class you wish to alter"
    Else
        Dim gxDataset As IGxDataset
        Set gxDataset = gxObject
      
        Dim dataset As IDataset
        Set dataset = gxDataset.dataset
        
        'make sure that now an actual dataset but rather a feature class is selected
        If (TypeOf dataset Is IFeatureDataset) Then
            MsgBox "Please select the individual feature class you wish to alter"
        ElseIf (TypeOf dataset Is IFeatureClass) Then
            Dim featureclass As IFeatureClass
            Set featureclass = dataset
            
            'make sure that the feature class does not already have an EXTCLSID value
            If ((featureclass.EXTCLSID Is Nothing)) Then
                Dim editSchema As IClassSchemaEdit
                Set editSchema = featureclass
                    
                Dim schemaLock As ISchemaLock
                Set schemaLock = featureclass
                          
                ' set an exclusive lock on the class
                schemaLock.ChangeSchemaLock (esriExclusiveSchemaLock)
                    
                ' create the IUID object
                Dim uid As New uid
                uid.Value = "EadaGDBExt.TimeStamper"
   
                ' Make the property set
                Dim propSet As IPropertySet
                Set propSet = New esriSystem.PropertySet
              
                propSet.SetProperty "CREATION_FIELDNAME", "Created"
                propSet.SetProperty "MODIFICATION_FIELDNAME", "ModifiedLast"
                propSet.SetProperty "USER_FIELDNAME", "ModifiedBy"
              
                ' alter the class extension for the class
                editSchema.AlterClassExtensionCLSID uid, propSet
   
                ' release the exclusive lock
                schemaLock.ChangeSchemaLock (esriSharedSchemaLock)
                    
                MsgBox "Altered feature class' EXTCLSID field"

            Else
                MsgBox "Cannot alter EXTCLSID for feature class. One already exists"
            End If
        End If
    End If
End If

End Sub




