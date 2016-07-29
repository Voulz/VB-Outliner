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
<Export(GetType(IMouseProcessorProvider))>
<ContentType("Basic")>
<TextViewRole(PredefinedTextViewRoles.Document)>
<Name("VBAdornment")>
Friend NotInheritable Class VBTaggerProvider
	Implements IWpfTextViewCreationListener, IMouseProcessorProvider

	<Import> Friend ClassificationRegistry As IClassificationTypeRegistryService = Nothing
	<Import> Friend Aggregator As IBufferTagAggregatorFactoryService = Nothing
	<Import> Friend ViewAgg As IViewTagAggregatorFactoryService
	<Import> Friend OutliningService As IOutliningManagerService = Nothing
	<Import> Friend ClassfifierService As IClassifierAggregatorService = Nothing

	Public Const LayerName As String = "VBOutline"

	<Export(GetType(AdornmentLayerDefinition))>
	<Name(LayerName)>
	<Order(After:=PredefinedAdornmentLayers.Text, Before:=PredefinedAdornmentLayers.Caret)>
	<Order(After:=PredefinedAdornmentLayers.Outlining)>
	<TextViewRole(PredefinedTextViewRoles.Document)>
	Public editorAdornmentLayer As AdornmentLayerDefinition


	Private Shared InLoop As Boolean = False

	Public Sub TextViewCreated(textView As IWpfTextView) Implements IWpfTextViewCreationListener.TextViewCreated

		GetAssociatedProcessor(textView)

	End Sub

	Public Function GetAssociatedProcessor(wpfTextView As IWpfTextView) As IMouseProcessor Implements IMouseProcessorProvider.GetAssociatedProcessor

		If InLoop Then Return Nothing

		InLoop = True
		Dim v = ViewAgg.CreateTagAggregator(Of IOutliningRegionTag)(wpfTextView)
		Dim o = OutliningService.GetOutliningManager(wpfTextView)
		Dim c = ClassfifierService.GetClassifier(wpfTextView.TextBuffer)
		InLoop = False

		Dim fi = FontAndColor.GetInstance()

		Return VBAdornment.Create(wpfTextView, v, o, c)
	End Function
End Class



Friend NotInheritable Class VBAdornment : Inherits MouseProcessorBase
	Public ReadOnly Keywords() As String = New String() {"Boolean", "Byte", "Char", "Integer",
			  "Date", "Decimal", "Double", "Integer", "Long", "Object", "SByte", "Short", "Single",
			  "String", "UInteger", "ULong", "UShort", "New", "Async"}

	Private ReadOnly _vaggregator As ITagAggregator(Of IOutliningRegionTag)
	Private ReadOnly _view As IWpfTextView
	Private ReadOnly _layer As IAdornmentLayer

	Private ReadOnly _OManager As IOutliningManager
	Private ReadOnly _classifier As IClassifier

	Private _tooltip As New ToolTip
	Private VisibleSpans As New Dictionary(Of ICollapsed, System.Windows.Rect)
	Private _lastToolTipOwner As ICollapsed = Nothing
	Private _tooltipTimer As Forms.Timer

	Private _changedTimer As Forms.Timer
	Private _changedSpan As New List(Of SnapshotSpan)

	Private NotNow As Boolean = False
	Public Shared Function Create(view As IWpfTextView, vaggregator As ITagAggregator(Of IOutliningRegionTag), OManager As IOutliningManager, Classifier As IClassifier) As VBAdornment

		Return view.Properties.GetOrCreateSingletonProperty(Of VBAdornment)(Function() New VBAdornment(view, vaggregator, OManager, Classifier))
	End Function
	Private Sub New(view As IWpfTextView, vaggregator As ITagAggregator(Of IOutliningRegionTag), OManager As IOutliningManager, Classifier As IClassifier)
		_view = view
		_vaggregator = vaggregator
		_layer = view.GetAdornmentLayer(VBTaggerProvider.LayerName)
		_OManager = OManager
		_classifier = Classifier

		AddHandler _OManager.RegionsCollapsed, AddressOf RegionsCollapsed
		AddHandler _view.TextBuffer.Changing, AddressOf BufferChanging
		AddHandler _view.LayoutChanged, AddressOf LayoutChanged
		AddHandler _view.Selection.SelectionChanged, AddressOf SelectionChanged
		AddHandler _view.Closed, AddressOf Closed
	End Sub
	Private Sub Closed(sender As Object, e As EventArgs)
		RemoveHandler _OManager.RegionsCollapsed, AddressOf RegionsCollapsed
		RemoveHandler _view.TextBuffer.Changing, AddressOf BufferChanging
		RemoveHandler _view.LayoutChanged, AddressOf LayoutChanged
		RemoveHandler _view.Selection.SelectionChanged, AddressOf SelectionChanged
		RemoveHandler _view.Closed, AddressOf Closed
	End Sub



	Private Sub BufferChanging(sender As Object, e As TextContentChangingEventArgs)
		If Not _changedSpan.Count = 0 Then NotNow = False : DrawLayout()
	End Sub

	Private Sub LayoutChanged(sender As Object, e As TextViewLayoutChangedEventArgs)


		Dim _spans = e.NewOrReformattedSpans
		If _spans.Count = 0 Then Return

		_changedSpan.AddRange(e.NewOrReformattedSpans)

		If NotNow Then Return
		DrawLayout()

	End Sub
	Private Sub DrawLayout()

		Try
			Dim _notDraw As New List(Of Microsoft.VisualStudio.Text.Outlining.ICollapsed)

			Dim _spans As New NormalizedSnapshotSpanCollection(_changedSpan)
			_changedSpan.Clear()

			If _spans.Count = 0 Then Return


			For Each t In _OManager.GetCollapsedRegions(_spans)
				_notDraw.AddRange(t.CollapsedChildren)

				Dim outline As New SnapshotSpan(_view.TextSnapshot, t.Extent.GetStartPoint(_view.TextSnapshot), t.CollapsedForm.ToString.Count)
				Dim _outG As Geometry = _view.TextViewLines.GetMarkerGeometry(outline)

				If _outG Is Nothing OrElse _notDraw.Contains(t) Then Continue For

				Dim start As String = t.CollapsedHintForm.ToString
				If start.StartsWith("#Region ") Then Continue For


				Dim bounds As System.Windows.Rect = _outG.Bounds
				Dim fi As FontAndColor = All.FontAndColor.GetInstance

				Dim bg As Brush = fi.BackgroundBrush
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

				Dim _collapsed As String = t.CollapsedForm.ToString
				'try to remove the ''' <summary> but if the comment is too short, the one behind is visible
				If _collapsed.StartsWith("'''") Then
					Dim l = _collapsed.Length
					_collapsed = _collapsed.Remove(0, "'''".Length).Trim
					If _collapsed.StartsWith("<summary>") Then _collapsed = _collapsed.Remove(0, "<summary>".Length).Trim
					_collapsed = _collapsed.PadLeft((l - _collapsed.Length) \ 2 + _collapsed.Length).PadRight(l)
				End If
				Dim txt As New FormattedText(_collapsed & " ", Globalization.CultureInfo.InvariantCulture, System.Windows.FlowDirection.LeftToRight, New Typeface(fi.FontName), fi.FontSize, fi.TextGreyBrush)
				bounds.Width = txt.WidthIncludingTrailingWhitespace + 2

				If start.StartsWith("'") Then
					txt.SetForegroundBrush(fi.CommentGreyBrush)
				Else
					For Each c In GetOutlineClassificationSpans(t)
						If c.ClassificationType.IsOfType(PredefinedClassificationTypeNames.Keyword) Then
							If Keywords.Contains(c.Text) Then
								txt.SetForegroundBrush(fi.KeywordBrush, c.Pos, c.Text.Length)
							Else
								txt.SetForegroundBrush(fi.OutlineBrush, c.Pos, c.Text.Length)
							End If
						ElseIf c.ClassificationType.IsOfType(PredefinedClassificationTypeNames.Identifier) Then
							If c.ClassificationType.IsOfType("VB User Types") Then 'dont hit in VS 2015
								txt.SetForegroundBrush(fi.UserTypeBrush, c.Pos, c.Text.Length)
							ElseIf c.ClassificationType.IsOfType("class name") Then 'dont hit either
								txt.SetForegroundBrush(fi.UserTypeClassesBrush, c.Pos, c.Text.Length)
							End If
							ElseIf c.ClassificationType.IsOfType(PredefinedClassificationTypeNames.String) Then
							txt.SetForegroundBrush(fi.StringBrush, c.Pos, c.Text.Length)
						ElseIf c.ClassificationType.IsOfType(PredefinedClassificationTypeNames.Number) Then
							txt.SetForegroundBrush(fi.NumberBrush, c.Pos, c.Text.Length)
						Else
						End If
					Next

					If txt.Text.EndsWith("... ") Then txt.SetForegroundBrush(fi.OutlineBrush, txt.Text.Length - 4, 4)
				End If

				dc.DrawRectangle(bg, New Pen(fi.OutlineBrush, 1), bounds)
				dc.DrawText(txt, New System.Windows.Point(bounds.X + 1, bounds.Y))
				dc.Close()
				group.Freeze()


				Dim drawingImage As New DrawingImage(group) : drawingImage.Freeze()
				Dim img As New Image() : img.Source = drawingImage

				If VisibleSpans.ContainsKey(t) Then VisibleSpans(t) = bounds Else VisibleSpans.Add(t, bounds)

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
		Dim Text As String = t.Extent.GetText(_view.TextSnapshot)
		Dim dec As Integer = 0

		For Each c In _classifier.GetClassificationSpans(t.Extent.GetSpan(_view.TextSnapshot))

			Dim iOut As Integer = TextOut.IndexOf(c.Span.GetText)

			If iOut = -1 Then
				'if we already found something, we leave, else, we continue searching
				If dec = 0 Then Continue For Else Exit For
			End If

			Dim i As Integer = Text.IndexOf(c.Span.GetText)
			Dim spantxt As String = c.Span.GetText

			w.Add(New word(spantxt, iOut + dec, c.ClassificationType))

			Text = Text.Remove(0, i + spantxt.Length)
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


	Public Overrides Sub PostprocessMouseLeave(e As Input.MouseEventArgs)
		MyBase.PostprocessMouseLeave(e)

		If Not _tooltip Is Nothing Then _tooltip.IsOpen = False
		If Not _tooltipTimer Is Nothing Then
			_tooltipTimer.Stop()
			RemoveHandler _tooltipTimer.Tick, AddressOf TooltipTick
			_tooltipTimer.Dispose()
			_tooltipTimer = Nothing
		End If
		_lastToolTipOwner = Nothing
	End Sub
	Public Overrides Sub PreprocessMouseMove(e As System.Windows.Input.MouseEventArgs)
		Dim pos = e.GetPosition(_view.VisualElement)
		pos.X += _view.ViewportLeft
		pos.Y += _view.ViewportTop

		Dim _notDraw As New List(Of Microsoft.VisualStudio.Text.Outlining.ICollapsed)


		Dim _span = _view.TextViewLines.FormattedSpan
		For Each t In _OManager.GetCollapsedRegions(_span)
			Dim _t As ICollapsed = t
			If Not t.CollapsedChildren.Count = 0 Then _notDraw.AddRange(t.CollapsedChildren)
			If _notDraw.Contains(t) OrElse Not VisibleSpans.ContainsKey(t) Then Continue For

			Dim bounds As System.Windows.Rect = VisibleSpans(t)
			If Not bounds.Contains(pos) Then Continue For
			e.Handled = True

			If _lastToolTipOwner Is t Then Exit For

			Dim fi As FontAndColor = All.FontAndColor.GetInstance

			_tooltip.IsOpen = False
			_tooltip.Background = fi.BackgroundBrush
			_tooltip.Foreground = fi.TextBrush
			_tooltip.BorderBrush = fi.OutlineBrush
			_lastToolTipOwner = t

			If Not _tooltipTimer Is Nothing Then
				RemoveHandler _tooltipTimer.Tick, AddressOf TooltipTick
				_tooltipTimer.Stop()
				_tooltipTimer.Dispose()
				_tooltipTimer = Nothing
			End If
			_tooltipTimer = New Forms.Timer
			_tooltipTimer.Interval = 500
			AddHandler _tooltipTimer.Tick, AddressOf TooltipTick
			_tooltipTimer.Tag = New Object() {_lastToolTipOwner, t}
			_tooltipTimer.Start()

			Exit For
		Next

		If Not e.Handled Then
			If Not _tooltip Is Nothing Then _tooltip.IsOpen = False
			If Not _tooltipTimer Is Nothing Then
				_tooltipTimer.Stop()
				RemoveHandler _tooltipTimer.Tick, AddressOf TooltipTick
				_tooltipTimer.Dispose()
				_tooltipTimer = Nothing
			End If
			_lastToolTipOwner = Nothing
		End If

	End Sub

	Private Sub TooltipTick(sender As Object, e As EventArgs)
		If _tooltipTimer IsNot Nothing Then
			If sender Is _tooltipTimer AndAlso _tooltip IsNot Nothing Then
				If _lastToolTipOwner Is DirectCast(sender, Forms.Timer).Tag(0) Then
					_tooltip.Content = ImproveHint(DirectCast(sender, Forms.Timer).Tag(1))
					_tooltip.IsOpen = True
				End If
			End If
			_tooltipTimer.Stop()
			RemoveHandler _tooltipTimer.Tick, AddressOf TooltipTick
			_tooltipTimer.Dispose()
			_tooltipTimer = Nothing
		End If
	End Sub
	Private Function ImproveHint(t As ICollapsed) As String

		Dim line = _view.TextSnapshot.GetLineFromPosition(t.Extent.GetStartPoint(_view.TextSnapshot))
		Dim nbSpaces As Integer = 0
		For Each c In line.GetText()
			If c = " "c Then nbSpaces += 1 Else Exit For
		Next

		Dim content As String = ""
		For Each l In t.CollapsedHintForm.ToString.Split(Environment.NewLine)
			If l(0) = " "c Then
				content &= Environment.NewLine & l.Remove(0, nbSpaces)
			ElseIf l.Length > 1 AndAlso Asc(l(0)) = 10 AndAlso l(1) = " "c Then
				content &= Environment.NewLine & l.Remove(0, nbSpaces + 1)
			Else
				content &= l
			End If
		Next
		Return content
	End Function


	Private Sub RegionsCollapsed(sender As Object, e As RegionsCollapsedEventArgs)

		NotNow = True

		If Not _changedTimer Is Nothing Then
			RemoveHandler _changedTimer.Tick, AddressOf ChangedTick
			_changedTimer.Stop()
			_changedTimer.Dispose()
			_changedTimer = Nothing
		End If

		_changedTimer = New Forms.Timer
		_changedTimer.Interval = 10
		AddHandler _changedTimer.Tick, AddressOf ChangedTick
		_changedTimer.Start()
	End Sub
	Private Sub ChangedTick(sender As Object, e As EventArgs)
		NotNow = False
		DrawLayout()
		_changedTimer.Stop()
	End Sub

	Private Sub SelectionChanged(sender As Object, e As EventArgs)
		'Dim tmp As String = ""

		'_changedSpan.AddRange(_view.Selection.SelectedSpans)

		'If NotNow Then Return
		'DrawLayout()
	End Sub
End Class
