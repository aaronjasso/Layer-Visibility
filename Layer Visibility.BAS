Sub Main
	'=================================================
	'Select layers to make visible. Mainly useful for the "Selected Layers On/Off"
	'option, which is not available in the model tree.
	'
	'Written by AMJ, 8.12.2019
	'=================================================
    Dim App as femap.model
    Set App = feFemap()

    '=================================================
	'Prompt user for layers to make visible
	'=================================================
    Dim layerSet As femap.Set
    Set layerSet = App.feSet
    Dim rc As Variant

    rc = layerSet.SelectMultiID(FT_LAYER,1,"Select Layers to Make Visible")

    If rc = 2 Then
    	Exit Sub
    End If

    '=================================================
	'Change view to "Show Visible Layers Only" (if necessary) and update layer
	'visibility.
	'=================================================
    Dim iterView As femap.View
    Set iterView = App.feView
    Dim vID As Long

	rc = App.feAppGetActiveView(vID)
    rc = iterView.Get(vID)
    If iterView.ShowAllLayers = True Then
    	iterView.ShowAllLayers = False
    End If

    Dim allLayers As femap.Set
    Set allLayers = App.feSet

    allLayers.AddAll(FT_LAYER)

    allLayers.Reset
    While allLayers.Next
    	If layerSet.IsAdded(allLayers.CurrentID) = -1 Then
    		iterView.SetLayerVisible(allLayers.CurrentID,True)
    	Else
			iterView.SetLayerVisible(allLayers.CurrentID,False)
    	End If
    Wend

	'Update view and regenerate to update window
	rc = iterView.Put(vID)
	App.feViewRegenerate(0)

    
End Sub
