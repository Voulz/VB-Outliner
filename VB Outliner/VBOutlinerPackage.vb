
Imports System.Diagnostics
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Drawing
'Imports EnvDTE
'Imports EnvDTE80
Imports Microsoft.VisualStudio
Imports Microsoft.VisualStudio.Shell.Interop
Imports Microsoft.VisualStudio.Shell
Imports System.Windows.Media

''' <summary>
''' This is the class that implements the package exposed by this assembly.
'''
''' The minimum requirement for a class to be considered a valid package for Visual Studio
''' is to implement the IVsPackage interface and register itself with the shell.
''' This package uses the helper classes defined inside the Managed Package Framework (MPF)
''' to do it: it derives from the Package class that provides the implementation of the 
''' IVsPackage interface and uses the registration attributes defined in the framework to 
''' register itself and its components with the shell.
''' </summary>
<PackageRegistration(UseManagedResourcesOnly:=True),
InstalledProductRegistration("#110", "#112", "1.0", IconResourceID:=400),
ProvideAutoLoad(UIContextGuids80.SolutionExists),
Guid(GuidList.guidVSPackagePkgString)>
Public NotInheritable Class VBOutlinerPackage
    Inherits Package

    ''' <summary>
    ''' Default constructor of the package.
    ''' Inside this method you can place any initialization code that does not require 
    ''' any Visual Studio service because at this point the package object is created but 
    ''' not sited yet inside Visual Studio environment. The place to do all the other 
    ''' initialization is the Initialize method.
    ''' </summary>
    Public Sub New()
        Debug.WriteLine(String.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", Me.GetType().Name))

    End Sub

#Region "Package Members"
    ''' <summary>
    ''' Initialization of the package; this method is called right after the package is sited, so this is the place
    ''' where you can put all the initialization code that rely on services provided by VisualStudio.
    ''' </summary>
    Protected Overrides Sub Initialize()
        Debug.WriteLine(String.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", Me.GetType().Name))
        MyBase.Initialize()

        Try
            FontAndColor.CreateInstance(GetService(GetType(IVsFontAndColorStorage)))
        Catch ex As Exception
        End Try

        'Try
        '	All.DTE = GetService(GetType(DTE))
        'Catch ex As Exception
        'End Try

        'Try
        '	All.DTE2 = System.Runtime.InteropServices.Marshal.GetActiveObject("VisualStudio.DTE.12.0")
        'Catch ex As Exception
        'End Try

    End Sub
#End Region
End Class

'Friend Module All
'    Public Property FontColorStorage As IVsFontAndColorStorage

Public Class FontAndColor
    Private Shared _instance As FontAndColor = Nothing
    Private Shared Property FontColorStorage As IVsFontAndColorStorage

    Private Sub New(FontColorStorage As IVsFontAndColorStorage)
        FontAndColor.FontColorStorage = FontColorStorage
        Try
            GetFontAndColor()
        Catch ex As Exception
        End Try
    End Sub
    Public Shared Function CreateInstance(FontColorStorage As IVsFontAndColorStorage) As FontAndColor
        If _instance Is Nothing Then
            _instance = New FontAndColor(FontColorStorage)
        End If
        Return _instance
    End Function
    Public Shared Function GetInstance() As FontAndColor
        Return _instance
    End Function
    Private Sub GetFontAndColor()


        'set default values before trying to read, if ever
        Me._FontName = "Consolas"
        Me._FontSize = 10

        Me._TextColor = System.Windows.Media.Color.FromRgb(220, 220, 220)
        Me._BackgroundColor = System.Windows.Media.Color.FromRgb(30, 30, 30)
        Me._SelectedColor = System.Windows.Media.Color.FromRgb(51, 153, 255)
        Me._InactiveColor = System.Windows.Media.Color.FromRgb(86, 86, 86)
        Me._CommentColor = System.Windows.Media.Color.FromRgb(87, 166, 74)
        Me._KeywordColor = System.Windows.Media.Color.FromRgb(86, 156, 214)
        Me._OutlineColor = System.Windows.Media.Color.FromRgb(128, 128, 128)
        Me._NumberColor = System.Windows.Media.Color.FromRgb(181, 206, 168)
        Me._StringColor = System.Windows.Media.Color.FromRgb(214, 157, 133)
        Me._UserTypeColor = System.Windows.Media.Color.FromRgb(78, 201, 176)

        If FontColorStorage Is Nothing Then Debug.Print("FontAndColor: FontColorStorage is nothing") : Return

        Dim c(0) As ColorableItemInfo
        Dim _guid = New System.Guid(FontsAndColorsCategory.TextEditor)
        Dim openresult As Integer = FontColorStorage.OpenCategory(_guid, __FCSTORAGEFLAGS.FCSF_LOADDEFAULTS Or __FCSTORAGEFLAGS.FCSF_READONLY Or __FCSTORAGEFLAGS.FCSF_NOAUTOCOLORS)
        If Not openresult = VSConstants.S_OK Then
            Debug.Print("FontAndColor: Not able to open FontColorStorage " & _guid.ToString)
        Else
            Try
                If FontColorStorage.GetItem("Plain Text", c) = VSConstants.S_OK Then
                    Me._TextColor = ConvertFromWin32Color(c(0).crForeground)
                    Me._BackgroundColor = ConvertFromWin32Color(c(0).crBackground)
                Else : Debug.Print("FontAndColor: Didn't find 'Plain Text'") : End If

                If FontColorStorage.GetItem("Selected Text", c) = VSConstants.S_OK Then
                    Me._SelectedColor = ConvertFromWin32Color(c(0).crBackground)
                Else : Debug.Print("FontAndColor: Didn't find 'Selected Text'") : End If

                If FontColorStorage.GetItem("Inactive Selected Text", c) = VSConstants.S_OK Then
                    Me._InactiveColor = ConvertFromWin32Color(c(0).crBackground)
                Else : Debug.Print("FontAndColor: Didn't find 'Inactive Selected Text'") : End If

                If FontColorStorage.GetItem("Comment", c) = VSConstants.S_OK Then
                    Me._CommentColor = ConvertFromWin32Color(c(0).crForeground)
                Else : Debug.Print("FontAndColor: Didn't find 'Comment'") : End If

                If FontColorStorage.GetItem("Keyword", c) = VSConstants.S_OK Then
                    Me._KeywordColor = ConvertFromWin32Color(c(0).crForeground)
                Else : Debug.Print("FontAndColor: Didn't find 'Keyword'") : End If

                If FontColorStorage.GetItem("Collapsible Text (Collapsed)", c) = VSConstants.S_OK Then
                    Me._OutlineColor = ConvertFromWin32Color(c(0).crForeground)
                Else : Debug.Print("FontAndColor: Didn't find 'Collapsible Text (Collapsed)'") : End If

                If FontColorStorage.GetItem("Number", c) = VSConstants.S_OK Then
                    Me._NumberColor = ConvertFromWin32Color(c(0).crForeground)
                Else : Debug.Print("FontAndColor: Didn't find 'Number'") : End If

                If FontColorStorage.GetItem("String", c) = VSConstants.S_OK Then
                    Me._StringColor = ConvertFromWin32Color(c(0).crForeground)
                Else : Debug.Print("FontAndColor: Didn't find 'String'") : End If

                Dim f(0) As LOGFONTW, fi(0) As FontInfo
                If FontColorStorage.GetFont(f, fi) = VSConstants.S_OK Then
                    Me._FontName = fi(0).bstrFaceName
                    Me._FontSize = Font.FromLogFont(f(0)).Size
                Else : Debug.Print("FontAndColor: Couldn't get the font") : End If

                If FontColorStorage.GetItem("User Types", c) = VSConstants.S_OK Then 'from previous version of VS
                    Me._UserTypeColor = ConvertFromWin32Color(c(0).crForeground)
                Else : Debug.Print("FontAndColor: Didn't find 'User Types'") : End If
            Catch ex As Exception
                Debug.Print("FontAndColor: Exception occured")
                Debug.Print(ex.Message)
            End Try

            FontColorStorage.CloseCategory()
        End If

        _guid = New System.Guid("{75A05685-00A8-4DED-BAE5-E7A50BFA929A}")
        openresult = FontColorStorage.OpenCategory(_guid, __FCSTORAGEFLAGS.FCSF_LOADDEFAULTS Or __FCSTORAGEFLAGS.FCSF_READONLY Or __FCSTORAGEFLAGS.FCSF_NOAUTOCOLORS)
        If Not openresult = VSConstants.S_OK Then
            Debug.Print("FontAndColor: Not able to open FontColorStorage " & _guid.ToString)
        Else
            Try
                If FontColorStorage.GetItem("class name", c) = VSConstants.S_OK Then
                    Me._UserTypeClassesColor = ConvertFromWin32Color(c(0).crForeground)
                Else : Debug.Print("FontAndColor: Didn't find 'class name'") : End If

                If FontColorStorage.GetItem("delegate name", c) = VSConstants.S_OK Then
                    Me._UserTypeDelegatesColor = ConvertFromWin32Color(c(0).crForeground)
                Else : Debug.Print("FontAndColor: Didn't find 'delegate name'") : End If

                If FontColorStorage.GetItem("enum name", c) = VSConstants.S_OK Then
                    Me._UserTypeEnumsColor = ConvertFromWin32Color(c(0).crForeground)
                Else : Debug.Print("FontAndColor: Didn't find 'enum name'") : End If

                If FontColorStorage.GetItem("interface name", c) = VSConstants.S_OK Then
                    Me._UserTypeInterfacesColor = ConvertFromWin32Color(c(0).crForeground)
                Else : Debug.Print("FontAndColor: Didn't find 'interface name'") : End If

                If FontColorStorage.GetItem("module name", c) = VSConstants.S_OK Then
                    Me._UserTypeModulesColor = ConvertFromWin32Color(c(0).crForeground)
                Else : Debug.Print("FontAndColor: Didn't find 'module name'") : End If

                If FontColorStorage.GetItem("struct name", c) = VSConstants.S_OK Then
                    Me._UserTypeStructuresColor = ConvertFromWin32Color(c(0).crForeground)
                Else : Debug.Print("FontAndColor: Didn't find 'struct name'") : End If

                If FontColorStorage.GetItem("type parameter name", c) = VSConstants.S_OK Then
                    Me._UserTypeTypeParametersColor = ConvertFromWin32Color(c(0).crForeground)
                Else : Debug.Print("FontAndColor: Didn't find 'type parameter name'") : End If

            Catch ex As Exception
                Debug.Print("FontAndColor: Exception occured")
                Debug.Print(ex.Message)
            End Try
            FontColorStorage.CloseCategory()
        End If

        Me._TextBrush = New SolidColorBrush(Me._TextColor)
        Me._SelectedBrush = New SolidColorBrush(Me._SelectedColor)
        Me._InactiveBrush = New SolidColorBrush(Me._InactiveColor)
        Me._BackgroundBrush = New SolidColorBrush(Me._BackgroundColor)
        Me._CommentBrush = New SolidColorBrush(Me._CommentColor)
        Me._KeywordBrush = New SolidColorBrush(Me._KeywordColor)
        Me._OutlineBrush = New SolidColorBrush(Me._OutlineColor)
        Me._NumberBrush = New SolidColorBrush(Me._NumberColor)
        Me._StringBrush = New SolidColorBrush(Me._StringColor)
        Me._UserTypeBrush = New SolidColorBrush(Me._UserTypeColor)

        Me._UserTypeClassesBrush = New SolidColorBrush(Me._UserTypeClassesColor)
        Me._UserTypeDelegatesBrush = New SolidColorBrush(Me._UserTypeDelegatesColor)
        Me._UserTypeEnumsBrush = New SolidColorBrush(Me._UserTypeEnumsColor)
        Me._UserTypeInterfacesBrush = New SolidColorBrush(Me._UserTypeInterfacesColor)
        Me._UserTypeModulesBrush = New SolidColorBrush(Me._UserTypeModulesColor)
        Me._UserTypeStructuresBrush = New SolidColorBrush(Me._UserTypeStructuresColor)
        Me._UserTypeTypeParametersBrush = New SolidColorBrush(Me._UserTypeTypeParametersColor)

        Me._KeywordGreyColor = MixColor(Me._KeywordColor, Me._OutlineColor, 0.4)
        Me._KeywordGreyBrush = New SolidColorBrush(Me._KeywordGreyColor)
        Me._CommentGreyColor = MixColor(Me._CommentColor, Me._OutlineColor, 0.4)
        Me._CommentGreyBrush = New SolidColorBrush(Me._CommentGreyColor)
        Me._TextGreyColor = MixColor(Me._TextColor, Me._OutlineColor, 0.6)
        Me._TextGreyBrush = New SolidColorBrush(Me._TextGreyColor)

    End Sub

    Public ReadOnly Property FontName As String
    Public ReadOnly Property FontSize As Double
    Public ReadOnly Property CommentColor As System.Windows.Media.Color
    Public ReadOnly Property KeywordColor As System.Windows.Media.Color
    Public ReadOnly Property TextColor As System.Windows.Media.Color
    Public ReadOnly Property TextGreyColor As System.Windows.Media.Color
    Public ReadOnly Property BackgroundColor As System.Windows.Media.Color
    Public ReadOnly Property OutlineColor As System.Windows.Media.Color
    Public ReadOnly Property KeywordGreyColor As System.Windows.Media.Color
    Public ReadOnly Property CommentGreyColor As System.Windows.Media.Color
    Public ReadOnly Property StringColor As System.Windows.Media.Color
    Public ReadOnly Property NumberColor As System.Windows.Media.Color
    Public ReadOnly Property SelectedColor As System.Windows.Media.Color
    Public ReadOnly Property InactiveColor As System.Windows.Media.Color
    Public ReadOnly Property UserTypeColor As System.Windows.Media.Color
    Public ReadOnly Property UserTypeClassesColor As System.Windows.Media.Color
    Public ReadOnly Property UserTypeDelegatesColor As System.Windows.Media.Color
    Public ReadOnly Property UserTypeEnumsColor As System.Windows.Media.Color
    Public ReadOnly Property UserTypeInterfacesColor As System.Windows.Media.Color
    Public ReadOnly Property UserTypeModulesColor As System.Windows.Media.Color
    Public ReadOnly Property UserTypeStructuresColor As System.Windows.Media.Color
    Public ReadOnly Property UserTypeTypeParametersColor As System.Windows.Media.Color


    Public ReadOnly Property CommentBrush As SolidColorBrush
    Public ReadOnly Property KeywordBrush As SolidColorBrush
    Public ReadOnly Property TextBrush As SolidColorBrush
    Public ReadOnly Property TextGreyBrush As SolidColorBrush
    Public ReadOnly Property BackgroundBrush As SolidColorBrush
    Public ReadOnly Property OutlineBrush As SolidColorBrush
    Public ReadOnly Property KeywordGreyBrush As SolidColorBrush
    Public ReadOnly Property CommentGreyBrush As SolidColorBrush
    Public ReadOnly Property StringBrush As SolidColorBrush
    Public ReadOnly Property NumberBrush As SolidColorBrush
    Public ReadOnly Property SelectedBrush As SolidColorBrush
    Public ReadOnly Property InactiveBrush As SolidColorBrush
    Public ReadOnly Property UserTypeBrush As SolidColorBrush
    Public ReadOnly Property UserTypeClassesBrush As SolidColorBrush
    Public ReadOnly Property UserTypeDelegatesBrush As SolidColorBrush
    Public ReadOnly Property UserTypeEnumsBrush As SolidColorBrush
    Public ReadOnly Property UserTypeInterfacesBrush As SolidColorBrush
    Public ReadOnly Property UserTypeModulesBrush As SolidColorBrush
    Public ReadOnly Property UserTypeStructuresBrush As SolidColorBrush
    Public ReadOnly Property UserTypeTypeParametersBrush As SolidColorBrush

    Public Shared Function ConvertFromWin32Color(color As Integer) As System.Windows.Media.Color
        Dim r As Integer = color And &HFF
        Dim g As Integer = (color And &HFF00) >> 8
        Dim b As Integer = (color And &HFF0000) >> 16
        Return New System.Windows.Media.Color With {.A = CByte(255), .R = CByte(r), .G = CByte(g), .B = CByte(b)}
    End Function
    Public Shared Function MixColor(c1 As System.Windows.Media.Color, c2 As System.Windows.Media.Color, Optional factor As Double = 0.5) As System.Windows.Media.Color

        factor = Math.Max(0, Math.Min(1, factor))
        Dim f As Double = 1D - factor
        Return New System.Windows.Media.Color With {.A = CByte(255),
         .R = (c1.R * factor + c2.R * f),
         .G = (c1.G * factor + c2.G * f),
         .B = (c1.B * factor + c2.B * f)}
    End Function
End Class
'End Module