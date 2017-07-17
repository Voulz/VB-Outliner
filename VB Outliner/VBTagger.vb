Imports System.ComponentModel.Composition
Imports Microsoft.VisualStudio.Text.Tagging
Imports Microsoft.VisualStudio.Utilities
Imports Microsoft.VisualStudio.Text
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports Microsoft.VisualStudio.Language.StandardClassification
Imports Microsoft.VisualStudio.Text.Classification
Imports Microsoft.VisualStudio.Text.Editor
Imports Microsoft.VisualStudio.Text.Outlining

<Export(GetType(IWpfTextViewCreationListener))>
<ContentType("Basic")>
<TextViewRole(PredefinedTextViewRoles.Document)>
<Name("VBAdornment")>
Friend NotInheritable Class VBTaggerProvider
	Implements IWpfTextViewCreationListener

	<Import> Friend OutliningService As IOutliningManagerService = Nothing
	<Import> Friend ViewClassifierService As IViewClassifierAggregatorService = Nothing

	Public Const LayerName As String = "VBOutline"

	<Export(GetType(AdornmentLayerDefinition))>
	<Name(LayerName)>
	<Order(After:=PredefinedAdornmentLayers.Text, Before:=PredefinedAdornmentLayers.Caret)>
	<Order(After:=PredefinedAdornmentLayers.Outlining)>
	<TextViewRole(PredefinedTextViewRoles.Document)>
	Public editorAdornmentLayer As AdornmentLayerDefinition


	Private Shared InLoop As Boolean = False

	Public Sub TextViewCreated(textView As IWpfTextView) Implements IWpfTextViewCreationListener.TextViewCreated
		If InLoop Then Return

		InLoop = True
		Dim o = OutliningService.GetOutliningManager(textView)
		Dim vc = ViewClassifierService.GetClassifier(textView)
		InLoop = False

		'FontAndColor.GetInstance()
		VBAdornment.Create(textView, o, vc)
	End Sub

End Class



Friend NotInheritable Class VBAdornment
	Public ReadOnly Keywords() As String = New String() {"Boolean", "Byte", "Char", "Integer",
			  "Date", "Decimal", "Double", "Integer", "Long", "Object", "SByte", "Short", "Single",
			  "String", "UInteger", "ULong", "UShort", "New", "Async"}

	Private ReadOnly _view As IWpfTextView
	Private ReadOnly _layer As IAdornmentLayer
	Private ReadOnly _OManager As IOutliningManager
	Private ReadOnly _viewClassifier As IClassifier

	Private _changedSpan As New List(Of SnapshotSpan)

	'Private NotNow As Boolean = False
	Public Shared Function Create(view As IWpfTextView, OManager As IOutliningManager, ViewClassifier As IClassifier) As VBAdornment
		Return view.Properties.GetOrCreateSingletonProperty(Of VBAdornment)(Function() New VBAdornment(view, OManager, ViewClassifier))
	End Function
	Private Sub New(view As IWpfTextView, OManager As IOutliningManager, ViewClassifier As IClassifier)
		_view = view
		_layer = view.GetAdornmentLayer(VBTaggerProvider.LayerName)
		_OManager = OManager
		_viewClassifier = ViewClassifier

		AddHandler _OManager.RegionsCollapsed, AddressOf RegionsCollapsed
		AddHandler _view.TextBuffer.Changing, AddressOf BufferChanging
		AddHandler _view.LayoutChanged, AddressOf LayoutChanged
		'AddHandler _view.Selection.SelectionChanged, AddressOf SelectionChanged
		AddHandler _view.Closed, AddressOf Closed
	End Sub
	Private Sub Closed(sender As Object, e As EventArgs)
		RemoveHandler _OManager.RegionsCollapsed, AddressOf RegionsCollapsed
		RemoveHandler _view.TextBuffer.Changing, AddressOf BufferChanging
		RemoveHandler _view.LayoutChanged, AddressOf LayoutChanged
		'RemoveHandler _view.Selection.SelectionChanged, AddressOf SelectionChanged
		RemoveHandler _view.Closed, AddressOf Closed
	End Sub



	Private Sub BufferChanging(sender As Object, e As TextContentChangingEventArgs)
		If Not _changedSpan.Count = 0 Then DrawLayout() 'NotNow = False :
	End Sub

	Private Sub LayoutChanged(sender As Object, e As TextViewLayoutChangedEventArgs)

		If e.NewOrReformattedSpans.Count > 0 Then
			_changedSpan.AddRange(e.NewOrReformattedSpans)
			'If Not NotNow Then
			DrawLayout()
		End If

	End Sub
	Private Sub DrawLayout()

		Try
			Dim _notDraw As New List(Of ICollapsed)

			Dim _spans As New NormalizedSnapshotSpanCollection(_changedSpan)
			_changedSpan.Clear()

			If _spans.Count = 0 Then Return


			For Each t In _OManager.GetCollapsedRegions(_spans)
				_notDraw.AddRange(t.CollapsedChildren)
				If _notDraw.Contains(t) Then Continue For

				Dim _collapsed As String = t.CollapsedForm.ToString
				Dim outline As New SnapshotSpan(_view.TextSnapshot, t.Extent.GetStartPoint(_view.TextSnapshot), _collapsed.Count)
				Dim _outG As Geometry = _view.TextViewLines.GetMarkerGeometry(outline)

				If _outG Is Nothing Then Continue For

				'Dim start As String = t.CollapsedHintForm.ToString
				If t.CollapsedHintForm.ToString.StartsWith("#Region ") Then Continue For


				Dim bounds As Rect = _outG.Bounds
				Dim _fontAndColors = FontAndColor.GetInstance

				Dim bg As Brush = _fontAndColors.BackgroundBrush
				'If Not _view.Selection.IsEmpty Then
				'    For Each s In _view.Selection.SelectedSpans
				'        If s.Contains(t.Extent.GetSpan(_view.TextSnapshot)) Then
				'            bg = If(_view.Selection.IsActive, fi.SelectedBrush, fi.InactiveBrush)
				'            Exit For
				'        End If
				'    Next
				'End If

				Dim group As New DrawingGroup
				Dim dc As DrawingContext = group.Open
				Dim isComment As Boolean = False
				' Dim _collapsed As String = t.CollapsedForm.ToString
				'try to remove the ''' <summary> but if the comment is too short, the one behind is visible
				If _collapsed.StartsWith("'''") Then
					Dim l = _collapsed.Length
					_collapsed = _collapsed.Remove(0, "'''".Length).Trim
					If _collapsed.StartsWith("<summary>") Then _collapsed = _collapsed.Remove(0, "<summary>".Length).Trim
					_collapsed = _collapsed.PadLeft((l - _collapsed.Length) \ 2 + _collapsed.Length).PadRight(l)
					isComment = True
				End If


				'Dim txt As New FormattedText(_collapsed & " ", Globalization.CultureInfo.InvariantCulture, System.Windows.FlowDirection.LeftToRight, New Typeface(fi.FontName), fi.FontSize, fi.TextGreyBrush)
				'TODO check PixelsPerDip
				Dim txt = New FormattedText(_collapsed & " ", Globalization.CultureInfo.InvariantCulture, FlowDirection.LeftToRight, New Typeface(_fontAndColors.FontName), _fontAndColors.FontSize, _fontAndColors.TextGreyBrush, 1)
				bounds.Width = txt.WidthIncludingTrailingWhitespace + 2

				If isComment OrElse _collapsed.StartsWith("'") Then
					txt.SetForegroundBrush(_fontAndColors.CommentGreyBrush)
				Else
					For Each c In GetOutlineClassificationSpans(t)
						If c.ClassificationType.IsOfType(PredefinedClassificationTypeNames.Keyword) Then
							If Keywords.Contains(c.Text) Then
								txt.SetForegroundBrush(_fontAndColors.KeywordBrush, c.Pos, c.Text.Length)
							Else
								txt.SetForegroundBrush(_fontAndColors.OutlineBrush, c.Pos, c.Text.Length)
							End If
						ElseIf c.ClassificationType.IsOfType(PredefinedClassificationTypeNames.String) Then
							txt.SetForegroundBrush(_fontAndColors.StringBrush, c.Pos, c.Text.Length)
						ElseIf c.ClassificationType.IsOfType(PredefinedClassificationTypeNames.Number) Then
							txt.SetForegroundBrush(_fontAndColors.NumberBrush, c.Pos, c.Text.Length)
						ElseIf c.ClassificationType.IsOfType("VB User Types") Then
							txt.SetForegroundBrush(_fontAndColors.UserTypeBrush, c.Pos, c.Text.Length)
						ElseIf c.ClassificationType.IsOfType("class name") Then
							txt.SetForegroundBrush(_fontAndColors.UserTypeClassesBrush, c.Pos, c.Text.Length)
						ElseIf c.ClassificationType.IsOfType("interface name") Then
							txt.SetForegroundBrush(_fontAndColors.UserTypeInterfacesBrush, c.Pos, c.Text.Length)
						ElseIf c.ClassificationType.IsOfType("delegate name") Then
							txt.SetForegroundBrush(_fontAndColors.UserTypeDelegatesBrush, c.Pos, c.Text.Length)
						ElseIf c.ClassificationType.IsOfType("enum name") Then
							txt.SetForegroundBrush(_fontAndColors.UserTypeEnumsBrush, c.Pos, c.Text.Length)
						ElseIf c.ClassificationType.IsOfType("module name") Then
							txt.SetForegroundBrush(_fontAndColors.UserTypeModulesBrush, c.Pos, c.Text.Length)
						ElseIf c.ClassificationType.IsOfType("struct name") Then
							txt.SetForegroundBrush(_fontAndColors.UserTypeStructuresBrush, c.Pos, c.Text.Length)
						ElseIf c.ClassificationType.IsOfType("type parameter name") Then
							txt.SetForegroundBrush(_fontAndColors.UserTypeTypeParametersBrush, c.Pos, c.Text.Length)
						End If
					Next

					If txt.Text.EndsWith("... ") Then txt.SetForegroundBrush(_fontAndColors.OutlineBrush, txt.Text.Length - 4, 4)
				End If

				dc.DrawRectangle(bg, New Pen(_fontAndColors.OutlineBrush, 1), bounds)
				dc.DrawText(txt, New System.Windows.Point(bounds.X + 1, bounds.Y))
				dc.Close()
				group.Freeze()


				Dim drawingImage As New DrawingImage(group) : drawingImage.Freeze()
				Dim img As New Image() With {.Source = drawingImage, .IsHitTestVisible = False}

				Canvas.SetLeft(img, _outG.Bounds.Left)
				Canvas.SetTop(img, _outG.Bounds.Top)

				_layer.AddAdornment(AdornmentPositioningBehavior.TextRelative, outline, Nothing, img, Nothing)
			Next
		Catch ex As Exception
		End Try
	End Sub

	Private Function GetOutlineClassificationSpans(t As ICollapsed) As List(Of word)

		Dim w As New List(Of word)
		Dim TextOut As String = t.CollapsedForm.ToString
		'Dim Text As String = t.Extent.GetText(_view.TextSnapshot)
		Dim dec As Integer = 0

#If DEBUG Then
		Dim allClassifiers As New Dictionary(Of String, ClassificationSpan)

		For Each c In _viewClassifier.GetClassificationSpans(t.Extent.GetSpan(_view.TextSnapshot))
			If Not allClassifiers.ContainsKey(c.ClassificationType.Classification) Then
				allClassifiers.Add(c.ClassificationType.Classification, c)
				Debug.Print(c.ClassificationType.Classification)
			End If
		Next
#End If

		For Each c In _viewClassifier.GetClassificationSpans(t.Extent.GetSpan(_view.TextSnapshot))
			Dim iOut As Integer = TextOut.IndexOf(c.Span.GetText)

			If iOut = -1 Then
				'if we already found something, we leave, else, we continue searching
				If dec = 0 Then Continue For Else Exit For
			End If

			Dim i As Integer = TextOut.IndexOf(c.Span.GetText) 'Text.IndexOf(c.Span.GetText)
			Dim spantxt As String = c.Span.GetText

			w.Add(New word(spantxt, iOut + dec, c.ClassificationType))

			'Text = Text.Remove(0, i + spantxt.Length)
			TextOut = TextOut.Remove(0, iOut + spantxt.Length)
			dec += iOut + spantxt.Length
		Next

		Return w
	End Function

	Private Structure word
		Public ReadOnly Text As String
		Public ReadOnly Pos As Integer
		Public ReadOnly ClassificationType As IClassificationType
		Public Sub New(t As String, p As Integer, c As IClassificationType)
			Me.Text = t
			Me.Pos = p
			Me.ClassificationType = c
		End Sub
	End Structure

	Private Sub RegionsCollapsed(sender As Object, e As RegionsCollapsedEventArgs)
		_view.VisualElement.InvalidateVisual()
	End Sub

	Private Sub SelectionChanged(sender As Object, e As EventArgs)
		'_changedSpan.AddRange(_view.Selection.SelectedSpans)

		'If NotNow Then Return
		'DrawLayout()
	End Sub
End Class
