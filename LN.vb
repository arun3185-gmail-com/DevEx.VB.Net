Imports System
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports Microsoft.VisualBasic.Interaction

Namespace LN
    
    ' Public Enum ACLLEVEL As Integer
    ' End Enum

    ' Public Enum ACLTYPE As Integer
    ' End Enum

    ' Public Enum AG_TARGET As Integer
    ' End Enum

    ' Public Enum AG_TRIGGER As Integer
    ' End Enum

    ' Public Enum DB_TYPES As Integer
    ' End Enum

' EMBED_TYPE
' FT_TYPES
' IT_TYPE
' LOG_EVENTS
' LOG_SEVERITY
' NOTES_ERRORS
' OE_CLASS
' OE_TYPE
' REG_TYPE
' RP_PRIORITY
' RT_ALIGN
' RT_EFFECTS
' RT_FONTS
' RT_PAGINATE
' RT_TAB
' SPACING
' USER_TYPE
' VC_ALIGN
' VC_DATEFMT
' VC_FONTSTYLE
' VC_NUMATTR
' VC_NUMFMT
' VC_SEP

    Public Enum NotesItemDataType As Integer
        UNKNOWN = 0
        RICHTEXT = 1
        COLLATION = 2
        NOTEREFS = 4
        ICON = 6
        NOTELINKS = 7
        SIGNATURE = 8
        USERDATA = 14
        QUERYCD = 15
        ACTIONCD = 16
        ASSISTANTINFO = 17
        VIEWMAPDATA = 18
        VIEWMAPLAYOUT = 19
        LSOBJECT = 20
        HTML = 21
        MIME_PART = 25
        ERRORITEM = 256
        UNAVAILABLE = 512
        NUMBERS = 768
        DATETIMES = 1024
        NAMES = 1074
        READERS = 1075
        AUTHORS = 1076
        ATTACHMENT = 1084
        OTHEROBJECT = 1085
        EMBEDDEDOBJECT = 1090
        TEXT = 1280
        TEXT_1281 = 1281
        FORMULA = 1536
        USERID = 1792
    End Enum

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Public Class NotesSession

        Private session As Object

        Public Sub New()
            session = CreateObject("Notes.NotesSession")
        End Sub

        Protected Overrides Sub Finalize()
            session = Nothing
        End Sub
        
        
        Public ReadOnly Property AddressBooks() As NotesDatabaseArray
            Get
                Return New LN.NotesDatabaseArray(session.AddressBooks)
            End Get
        End Property

        Public ReadOnly Property CommonUserName() As String
            Get
                Return session.CommonUserName
            End Get
        End Property

        Public ReadOnly Property ConvertMIME() As String
            Get
                Return session.ConvertMIME
            End Get
        End Property

        Public ReadOnly Property CurrentDatabase() As NotesDatabase
            Get
                Return New LN.NotesDatabase(session.CurrentDatabase)
            End Get
        End Property

        Public ReadOnly Property DocumentContext() As NotesDocument
            Get
                Return New LN.NotesDocument(session.DocumentContext)
            End Get
        End Property

        Public ReadOnly Property EffectiveUserName() As String
            Get
                Return session.EffectiveUserName
            End Get
        End Property

        Public ReadOnly Property HttpURL() As String
            Get
                Return session.HttpURL
            End Get
        End Property

        Public ReadOnly Property IsOnServer() As Boolean
            Get
                Return session.IsOnServer
            End Get
        End Property

        Public ReadOnly Property LastExitStatus() As Long
            Get
                Return session.LastExitStatus
            End Get
        End Property

        ' Public ReadOnly Property LastRun() As TTTTT
        '     Get
        '         Return session.LastRun
        '     End Get
        ' End Property

        Public ReadOnly Property NotesBuildVersion() As Long
            Get
                Return session.NotesBuildVersion
            End Get
        End Property

        Public ReadOnly Property NotesURL() As String
            Get
                Return session.NotesURL
            End Get
        End Property

        Public ReadOnly Property NotesVersion() As String
            Get
                Return session.NotesVersion
            End Get
        End Property

        Public ReadOnly Property OrgDirectoryPath() As String
            Get
                Return session.OrgDirectoryPath
            End Get
        End Property

        Public ReadOnly Property Platform() As String
            Get
                Return session.Platform
            End Get
        End Property

        Public ReadOnly Property SavedData() As NotesDocument
            Get
                Return New LN.NotesDocument(session.SavedData)
            End Get
        End Property

        Public ReadOnly Property ServerName() As String
            Get
                Return session.ServerName
            End Get
        End Property

        Public ReadOnly Property URLDatabase() As NotesDatabase
            Get
                Return New LN.NotesDatabase(session.URLDatabase)
            End Get
        End Property

        ' Public ReadOnly Property xxxxx() As TTTTTT
        '     Get
        '         Return session.xxxxx
        '     End Get
        ' End Property
        
        Public ReadOnly Property UserName() As String
            Get
                Return session.UserName
            End Get
        End Property

        Public Function CreateDateTime(ByVal dateTime As String) As NotesDateTime
            Return New LN.NotesDateTime(session.CreateDateTime(dateTime))
        End Function

        Public Function Evaluate(ByVal formula As String, ByVal doc As LN.NotesDocument) As Object
            Return session.Evaluate(formula, doc)
        End Function

        Public Function GetDatabase(ByVal serverName As String, ByVal lnFilePath As String) As NotesDatabase
            Return New LN.NotesDatabase(session.GetDatabase(serverName, lnFilePath))
        End Function

        Public Function GetEnvironmentString(ByVal name As String, Optional ByVal system As Boolean = False) As Object
            Return session.GetEnvironmentString(name, system)
        End Function

        Public Function GetEnvironmentValue(ByVal name As String, Optional ByVal system As Boolean = False) As Object
            Return session.GetEnvironmentValue(name, system)
        End Function

        Public Function SendConsoleCommand(ByVal serverName As String, ByVal consoleCommand As String) As String
            Return session.SendConsoleCommand(serverName, consoleCommand)
        End Function

    End Class


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Public Class NotesDatabase
        
        Private Database As Object

        Public Sub New(ByRef database As Object)
            Me.Database = database
        End Sub

        Protected Overrides Sub Finalize()
            Database = Nothing
        End Sub


        Public ReadOnly Property AllDocuments() As NotesDocumentCollection
            Get
                Return New LN.NotesDocumentCollection(Database.AllDocuments)
            End Get
        End Property

        Public ReadOnly Property Categories() As String
            Get
                Return Database.Categories
            End Get
        End Property

        Public ReadOnly Property Created() As DateTime
            Get
                Return Database.Created
            End Get
        End Property
        
        Public ReadOnly Property CurrentAccessLevel() As Integer
            Get
                Return Database.CurrentAccessLevel
            End Get
        End Property

        Public ReadOnly Property DesignTemplateName() As String
            Get
                Return Database.DesignTemplateName
            End Get
        End Property

        Public ReadOnly Property FileFormat() As Integer
            Get
                Return Database.FileFormat
            End Get
        End Property

        Public ReadOnly Property FileName() As String
            Get
                Return Database.FileName
            End Get
        End Property

        Public ReadOnly Property FilePath() As String
            Get
                Return Database.FilePath
            End Get
        End Property

        Public ReadOnly Property FolderReferencesEnabled() As Boolean
            Get
                Return Database.FolderReferencesEnabled
            End Get
        End Property

        Public ReadOnly Property Forms() As NotesFormCollection
            Get
                Return New LN.NotesFormCollection(Database.Forms)
            End Get
        End Property

        Public ReadOnly Property HttpURL() As String
            Get
                Return Database.HttpURL
            End Get
        End Property

        Public ReadOnly Property LastModified() As DateTime
            Get
                Return Database.LastModified
            End Get
        End Property

        Public ReadOnly Property NotesURL() As String
            Get
                Return Database.NotesURL
            End Get
        End Property

        Public ReadOnly Property ReplicaID() As String
            Get
                Return Database.ReplicaID
            End Get
        End Property

        Public ReadOnly Property Server() As String
            Get
                Return Database.Server
            End Get
        End Property

        Public ReadOnly Property Size() As Double
            Get
                Return Database.Size
            End Get
        End Property

        Public ReadOnly Property TemplateName() As String
            Get
                Return Database.TemplateName
            End Get
        End Property

        Public ReadOnly Property Title() As String
            Get
                Return Database.Title
            End Get
        End Property

        Public ReadOnly Property Views() As LN.NotesView()
            Get
                Return Database.Views
            End Get
        End Property


        Public Function GetDocumentByID(ByVal NoteID As String) As LN.NotesDocument
            Return New LN.NotesDocument(Database.GetDocumentByID(NoteID))
        End Function

        Public Function GetDocumentByUNID(ByVal Unid As String) As LN.NotesDocument
            Return New LN.NotesDocument(Database.GetDocumentByUNID(Unid))
        End Function

    End Class
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Public Class NotesDatabaseArray

        Private DatabaseArray As Object
        
        Public Sub New(ByRef databaseArray As Object)
            Me.DatabaseArray = databaseArray
        End Sub

        Default Public ReadOnly Property Item(ByVal Index As Integer) As NotesDatabase
            Get
                Return New LN.NotesDatabase(Me.DatabaseArray(Index))
            End Get
        End Property

        Public ReadOnly Property Length() As Integer
            Get
                Return Me.DatabaseArray.Length
            End Get
        End Property

        Protected Overrides Sub Finalize()
            DatabaseArray = Nothing
        End Sub

    End Class

    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Public Class NotesView
    
        Private nView As Object
        
        Public Sub New(ByRef nView As Object)
            Me.nView = nView
        End Sub
        
        Protected Overrides Sub Finalize()
            Me.nView = Nothing
        End Sub

        Public ReadOnly Property Aliases() As String()
            Get
                Return nView.Aliases
            End Get
        End Property
        
        Public ReadOnly Property AutoUpdate() As Boolean
            Get
                Return nView.AutoUpdate
            End Get
        End Property

        Public ReadOnly Property BackgroundColor() As Integer
            Get
                Return nView.BackgroundColor
            End Get
        End Property

        Public ReadOnly Property ColumnCount() As Integer
            Get
                Return nView.ColumnCount
            End Get
        End Property

        Public ReadOnly Property ColumnNames() As Object
            Get
                Return nView.ColumnNames
            End Get
        End Property

        'Public ReadOnly Property Columns() As NotesViewColumn()
        '    Get
        '        Return nView.Columns
        '    End Get
        'End Property

        Public ReadOnly Property Created() As Date
            Get
                Return nView.Created
            End Get
        End Property

        Public ReadOnly Property EntryCount() As Integer
            Get
                Return nView.EntryCount
            End Get
        End Property

        Public ReadOnly Property HttpURL() As String
            Get
                Return nView.HttpURL
            End Get
        End Property

        Public ReadOnly Property IsCalendar() As Boolean
            Get
                Return nView.IsCalendar
            End Get
        End Property

        Public ReadOnly Property IsCategorized() As Boolean
            Get
                Return nView.IsCategorized
            End Get
        End Property

        Public ReadOnly Property IsConflict() As Boolean
            Get
                Return nView.IsConflict
            End Get
        End Property

        Public ReadOnly Property IsDefaultView() As Boolean
            Get
                Return nView.IsDefaultView
            End Get
        End Property

        Public ReadOnly Property IsFolder() As Boolean
            Get
                Return nView.IsFolder
            End Get
        End Property

        Public ReadOnly Property IsHierarchical() As Boolean
            Get
                Return nView.IsHierarchical
            End Get
        End Property

        Public ReadOnly Property IsModified() As Boolean
            Get
                Return nView.IsModified
            End Get
        End Property

        Public ReadOnly Property IsPrivate() As Boolean
            Get
                Return nView.IsPrivate
            End Get
        End Property

        Public ReadOnly Property IsProhibitDesignRefresh() As Boolean
            Get
                Return nView.IsProhibitDesignRefresh
            End Get
        End Property

        Public ReadOnly Property LastModified() As Date
            Get
                Return nView.LastModified
            End Get
        End Property

        Public ReadOnly Property LockHolders() As String()
            Get
                Return nView.LockHolders
            End Get
        End Property

        Public ReadOnly Property Name() As String
            Get
                Return nView.Name
            End Get
        End Property

        Public ReadOnly Property NotesURL() As String
            Get
                Return nView.NotesURL
            End Get
        End Property

        Public ReadOnly Property ProtectReaders() As Boolean
            Get
                Return nView.ProtectReaders
            End Get
        End Property

        Public ReadOnly Property Readers() As String()
            Get
                Return nView.Readers
            End Get
        End Property

        Public ReadOnly Property RowLines() As Integer
            Get
                Return nView.RowLines
            End Get
        End Property

        Public ReadOnly Property SelectionFormula() As String
            Get
                Return nView.SelectionFormula
            End Get
        End Property

        Public ReadOnly Property Spacing() As Integer
            Get
                Return nView.Spacing
            End Get
        End Property

        Public ReadOnly Property TopLevelEntryCount() As Integer
            Get
                Return nView.TopLevelEntryCount
            End Get
        End Property

        Public ReadOnly Property UniversalID() As String
            Get
                Return nView.UniversalID
            End Get
        End Property

        Public ReadOnly Property ViewInheritedName() As String
            Get
                Return nView.ViewInheritedName
            End Get
        End Property

        Public Sub Clear()
            nView.Clear()
        End Sub

        'Public Function CreateViewNav(ByVal cacheSize As Long) As LN.NotesViewNavigator
        '    Return nView.CreateViewNav(cacheSize)
        'End Function

        'Public Function CreateViewNavFrom(ByRef navigatorObject As Object, Optional ByVal cacheSize As Long = 0) As LN.NotesViewNavigator
        '    Return nView.CreateViewNavFrom(navigatorObject, cacheSize)
        'End Function
        
        'Public Function CreateViewNavFromCategory(ByVal category As String, Optional ByVal cacheSize As Long = 0) As LN.NotesViewNavigator
        '    Return nView.CreateViewNavFromCategory(category, cacheSize)
        'End Function

        'Public Function CreateViewNavFromChildren(ByRef navigatorObject As Object, Optional ByVal cacheSize As Long = 0) As LN.NotesViewNavigator
        '    Return nView.CreateViewNavFromChildren(navigatorObject, cacheSize)
        'End Function

        'Public Function CreateViewNavFromDescendants(ByRef navigatorObject As Object, Optional ByVal cacheSize As Long = 0) As LN.NotesViewNavigator
        '    Return nView.CreateViewNavFromDescendants(navigatorObject, cacheSize)
        'End Function

        'Public Function CreateViewNavMaxLevel(ByVal level As Long, Optional ByVal cacheSize As Long = 0) As LN.NotesViewNavigator
        '    Return nView.CreateViewNavMaxLevel(level, cacheSize)
        'End Function

        Public Function FTSearch(ByVal query As String, ByVal maxDocs As Integer) As Long
            Return nView.FTSearch(query, maxDocs)
        End Function

        Public Function GetAllDocumentsByKey(ByVal keyArray As String, Optional ByVal exactMatch As Boolean = False) As LN.NotesDocumentCollection
            Return New LN.NotesDocumentCollection(nView.GetAllDocumentsByKey(keyArray, exactMatch))
        End Function

        'Public Function GetAllEntriesByKey(ByVal keyArray As String, Optional ByVal exactMatch As Boolean = False) As LN.NotesViewEntryCollection
        '    Return nView.GetAllEntriesByKey(keyArray, exactMatch)
        'End Function

        Public Function GetChild(ByRef document As LN.NotesDocument) As LN.NotesDocument
            Return New LN.NotesDocument(nView.GetChild(document))
        End Function

        'Public Function GetColumn(ByVal columnNumber As Long) As LN.NotesViewColumn
        '    Return New LN.NotesViewColumn(nView.GetColumn(columnNumber))
        'End Function
        
        Public Function GetDocumentByKey(ByVal keyArray As Long, ByVal exactMatch As Boolean) As LN.NotesDocument
            Return New LN.NotesDocument(nView.GetDocumentByKey(keyArray, exactMatch))
        End Function

        'Public Function GetEntryByKey(ByVal keyArray As String, Optional ByVal exactMatch As Boolean = False) As LN.NotesViewEntry
        '    Return nView.GetEntryByKey(keyArray, exactMatch)
        'End Function

        Public Function GetFirstDocument() As LN.NotesDocument
            Return New LN.NotesDocument(nView.GetFirstDocument())
        End Function

        Public Function GetLastDocument() As LN.NotesDocument
            Return New LN.NotesDocument(nView.GetLastDocument())
        End Function

        Public Function GetNextDocument(ByRef document As LN.NotesDocument) As LN.NotesDocument
            Return New LN.NotesDocument(nView.GetNextDocument(document))
        End Function

        Public Function GetNextSibling(ByRef document As LN.NotesDocument) As LN.NotesDocument
            Return New LN.NotesDocument(nView.GetNextSibling(document))
        End Function

        Public Function GetNthDocument(ByVal index As Long) As LN.NotesDocument
            Return New LN.NotesDocument(nView.GetNthDocument(index))
        End Function

        Public Function GetParentDocument(ByRef document As LN.NotesDocument) As LN.NotesDocument
            Return New LN.NotesDocument(nView.GetParentDocument(document))
        End Function

        Public Function GetPrevDocument(ByRef document As LN.NotesDocument) As LN.NotesDocument
            Return New LN.NotesDocument(nView.GetPrevDocument(document))
        End Function

        Public Function GetPrevSibling(ByRef document As LN.NotesDocument) As LN.NotesDocument
            Return New LN.NotesDocument(nView.GetPrevSibling(document))
        End Function

        Public Sub Refresh()
            nView.Refresh()
        End Sub

    End Class
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Public Class NotesViewArray
        
        Private nViews As Object

        Public Sub New(ByRef nViews As Object)
            Me.nViews = nViews
        End Sub

        Public ReadOnly Property Count() As Long
            Get
                Return Me.nViews.Count
            End Get
        End Property

        Default Public ReadOnly Property Item(ByVal Index As Integer) As NotesView
            Get
                Return New LN.NotesView(Me.nViews(Index))
            End Get
        End Property

        Public ReadOnly Property Length() As Integer
            Get
                Return Me.nViews.Length
            End Get
        End Property
        
        Protected Overrides Sub Finalize()
            Me.nViews = Nothing
        End Sub

    End Class
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Public Class NotesForm

        Private nForm As Object

        Public Sub New(ByRef nForm As Object)
            Me.nForm = nForm
        End Sub

        Protected Overrides Sub Finalize()
            Me.nForm = Nothing
        End Sub

        Public ReadOnly Property Aliases() As String()
            Get
                Return nForm.Aliases
            End Get
        End Property
        
        Public ReadOnly Property Fields() As String()
            Get
                Return nForm.Fields
            End Get
        End Property

        Public ReadOnly Property FormUsers() As String()
            Get
                Return nForm.FormUsers
            End Get
        End Property

        Public ReadOnly Property HttpURL() As String
            Get
                Return nForm.HttpURL
            End Get
        End Property
        
        Public ReadOnly Property IsSubForm() As Boolean
            Get
                Return nForm.IsSubForm
            End Get
        End Property
        
        Public ReadOnly Property LockHolders() As String()
            Get
                Return nForm.LockHolders
            End Get
        End Property

        Public ReadOnly Property Name() As String
            Get
                Return nForm.Name
            End Get
        End Property

        Public ReadOnly Property NotesURL() As String
            Get
                Return nForm.NotesURL
            End Get
        End Property

        Public ReadOnly Property Readers() As String()
            Get
                Return nForm.Readers
            End Get
        End Property

        Public Function GetFieldType(ByVal name As String) As Integer
            Return nForm.GetFieldType(name)
        End Function

    End Class
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Public Class NotesFormCollection

        Private FormCollection As Object
        
        Public Sub New(ByRef formCollection As Object)
            Me.FormCollection = formCollection
        End Sub

        Default Public ReadOnly Property Item(ByVal Index As Integer) As NotesForm
            Get
                Return New LN.NotesForm(Me.FormCollection(Index))
            End Get
        End Property

        Public ReadOnly Property Length() As Integer
            Get
                Return Me.FormCollection.Length
            End Get
        End Property

        Protected Overrides Sub Finalize()
            FormCollection = Nothing
        End Sub

    End Class
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Public Class NotesDocument
        
        Friend Doc As Object

        Public Sub New(ByRef doc As Object)
            Me.Doc = doc
        End Sub

        Protected Overrides Sub Finalize()
            Doc = Nothing
        End Sub


        Public ReadOnly Property Authors() As String()
            Get
                Return Doc.Authors
            End Get
        End Property

        Public ReadOnly Property ColumnValues() As Object()
            Get
                Return Doc.ColumnValues
            End Get
        End Property

        Public ReadOnly Property Created() As DateTime
            Get
                Return Doc.Created
            End Get
        End Property

        Public ReadOnly Property HasEmbedded() As Boolean
            Get
                Return Doc.HasEmbedded
            End Get
        End Property

        Public ReadOnly Property HttpURL() As String
            Get
                Return Doc.HttpURL
            End Get
        End Property

        Public ReadOnly Property IsDeleted() As Boolean
            Get
                Return Doc.IsDeleted
            End Get
        End Property

        Public ReadOnly Property IsEncrypted() As Boolean
            Get
                Return Doc.IsEncrypted
            End Get
        End Property

        Public ReadOnly Property IsNewNote() As Boolean
            Get
                Return Doc.IsNewNote
            End Get
        End Property

        Public ReadOnly Property IsProfile() As Boolean
            Get
                Return Doc.IsProfile
            End Get
        End Property

        Public ReadOnly Property IsResponse() As Boolean
            Get
                Return Doc.IsResponse
            End Get
        End Property

        Public ReadOnly Property IsSigned() As Boolean
            Get
                Return Doc.IsSigned
            End Get
        End Property

        Public ReadOnly Property IsUIDocOpen() As Boolean
            Get
                Return Doc.IsUIDocOpen
            End Get
        End Property

        Public ReadOnly Property IsValid() As Boolean
            Get
                Return Doc.IsValid
            End Get
        End Property

        Public ReadOnly Property Items() As NotesItemArray
           Get
               Return New NotesItemArray(Doc.Items)
           End Get
        End Property

        Public ReadOnly Property Key() As String
            Get
                Return Doc.Key
            End Get
        End Property

        Public ReadOnly Property LastAccessed() As DateTime
            Get
                Return Doc.LastAccessed
            End Get
        End Property

        Public ReadOnly Property LastModified() As DateTime
            Get
                Return Doc.LastModified
            End Get
        End Property

        Public ReadOnly Property NoteID() As String
            Get
                Return Doc.NoteID
            End Get
        End Property

        Public ReadOnly Property NotesURL() As String
            Get
                Return Doc.NotesURL
            End Get
        End Property

        Public ReadOnly Property Size() As Long
            Get
                Return Doc.Size
            End Get
        End Property

        Public ReadOnly Property UniversalID() As String
            Get
                Return Doc.UniversalID
            End Get
        End Property

        Public Function GetFirstItem(ByVal name As String) As NotesItem
            Return New NotesItem(Doc.GetFirstItem(name))
        End Function

        Public Function GetItemValue(ByVal itemName As String) As System.Array
            Return Doc.GetItemValue(itemName)
        End Function

        Public Function GetItemValueCustomDataBytes(ByVal itemName As String, ByVal dataTypeName As String) As Byte()
            Return Doc.GetItemValueCustomDataBytes(itemName, dataTypeName)
        End Function

        Public Function HasItem(ByVal itemName As String) As Boolean
            Return Doc.HasItem(itemName)
        End Function

    End Class
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Public Class NotesDocumentCollection
        
        Private DocCollection As Object
        
        Public Sub New(ByRef docCollection As Object)
            Me.DocCollection = docCollection
        End Sub

        Protected Overrides Sub Finalize()
            DocCollection = Nothing
        End Sub


        Public ReadOnly Property Count() As Long
            Get
                Return DocCollection.Count
            End Get
        End Property

        Public ReadOnly Property IsSorted() As Boolean
            Get
                Return DocCollection.IsSorted
            End Get
        End Property

        Public ReadOnly Property Query() As String
            Get
                Return DocCollection.Query
            End Get
        End Property


        Public Function GetFirstDocument() As LN.NotesDocument
            Dim doc1 As Object = DocCollection.GetFirstDocument()
            If doc1 Is Nothing Then
                Return Nothing
            Else
                Return New LN.NotesDocument(doc1)
            End If
        End Function

        Public Function GetPrevDocument(doc As LN.NotesDocument) As LN.NotesDocument
            Dim doc1 As Object = DocCollection.GetPrevDocument(doc.Doc)
            If doc1 Is Nothing Then
                Return Nothing
            Else
                Return New LN.NotesDocument(doc1)
            End If
        End Function

        Public Function GetNextDocument(doc As LN.NotesDocument) As LN.NotesDocument
            Dim doc1 As Object = DocCollection.GetNextDocument(doc.Doc)
            If doc1 Is Nothing Then
                Return Nothing
            Else
                Return New LN.NotesDocument(doc1)
            End If
        End Function

        Public Function GetLastDocument() As LN.NotesDocument
            Dim doc1 As Object = DocCollection.GetLastDocument()
            If doc1 Is Nothing Then
                Return Nothing
            Else
                Return New LN.NotesDocument(doc1)
            End If
        End Function

    End Class
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Public Class NotesItem
        
        Protected nItem As Object

        Public Sub New(ByRef item As Object)
            Me.nItem = item
        End Sub

        'Public Sub New(ByRef notesDocument As LN.NotesDocument, ByVal name As String, ByVal value As Object, Optional ByVal specialType As Integer)
        '    Me.nItem = item
        'End Sub

        Protected Overrides Sub Finalize()
            nItem = Nothing
        End Sub

        Public ReadOnly Property DateTimeValue() As NotesDateTime
            Get
                Return New LN.NotesDateTime(nItem.DateTimeValue)
            End Get
        End Property

        Public ReadOnly Property IsAuthors() As Boolean
            Get
                Return nItem.IsAuthors
            End Get
        End Property

        Public ReadOnly Property IsEncrypted() As Boolean
            Get
                Return nItem.IsEncrypted
            End Get
        End Property

        Public ReadOnly Property IsNames() As Boolean
            Get
                Return nItem.IsNames
            End Get
        End Property

        Public ReadOnly Property IsProtected() As Boolean
            Get
                Return nItem.IsProtected
            End Get
        End Property

        Public ReadOnly Property IsReaders() As Boolean
            Get
                Return nItem.IsReaders
            End Get
        End Property

        Public ReadOnly Property IsSigned() As Boolean
            Get
                Return nItem.IsSigned
            End Get
        End Property

        Public ReadOnly Property IsSummary() As Boolean
            Get
                Return nItem.IsSummary
            End Get
        End Property

        Public ReadOnly Property LastModified() As Date
            Get
                Return nItem.LastModified
            End Get
        End Property

        Public ReadOnly Property Name() As String
            Get
                Return nItem.Name
            End Get
        End Property

        Public ReadOnly Property SaveToDisk() As Boolean
            Get
                Return nItem.SaveToDisk
            End Get
        End Property

        Public ReadOnly Property Text() As String
            Get
                Return nItem.Text
            End Get
        End Property

        Public ReadOnly Property Type() As Long
            Get
                Return nItem.Type
            End Get
        End Property

        Public ReadOnly Property ValueLength() As Long
            Get
                Return nItem.ValueLength
            End Get
        End Property
        
        
        Public Function GetValueCustomDataBytes(ByVal dataTypeName As String) As Byte()
            Return nItem.GetValueCustomDataBytes(dataTypeName)
        End Function


    End Class
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Public Class NotesItemArray
        
        Private items As Object

        Public Sub New(ByRef items As Object)
            Me.items = items
        End Sub

        Default Public ReadOnly Property Item(ByVal Index As Integer) As LN.NotesItem
            Get
                Return New LN.NotesItem(Me.items(Index))
            End Get
        End Property

        Public ReadOnly Property Length() As Integer
            Get
                Return Me.items.Length
            End Get
        End Property

        Protected Overrides Sub Finalize()
            Me.items = Nothing
        End Sub

    End Class
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Public Class NotesRichTextItem
        Inherits NotesItem
        
        Public Sub New(ByRef item As Object)
            MyBase.New(item)
        End Sub

        Protected Overrides Sub Finalize()
            MyBase.nItem = Nothing
        End Sub

        
        Public Function GetFormattedText(ByVal tabstrip As Boolean, ByVal lineLength As Integer) As String
            Return MyBase.nItem.GetFormattedText()
        End Function

        Public Function GetUnformattedText() As String
            Return MyBase.nItem.GetUnformattedText()
        End Function

    End Class
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' Public Class NotesName
    ' End Class

    ' Public Class NotesAgent
    ' End Class

    Public Class NotesDateTime

        Private notesDateTime As Object
        
        Public Sub New(ByVal notesDateTime As Object)
            Me.notesDateTime = notesDateTime
        End Sub
        
        Protected Overrides Sub Finalize()
            Me.notesDateTime = Nothing
        End Sub

        Public ReadOnly Property DateOnly() As String
            Get
                Return notesDateTime.DateOnly
            End Get
        End Property

        Public ReadOnly Property GMTTime() As String
            Get
                Return notesDateTime.GMTTime
            End Get
        End Property

        Public ReadOnly Property IsDST() As Boolean
            Get
                Return notesDateTime.IsDST
            End Get
        End Property

        Public ReadOnly Property IsValidDate() As Boolean
            Get
                Return notesDateTime.IsValidDate
            End Get
        End Property

        Public ReadOnly Property LocalTime() As String
            Get
                Return notesDateTime.LocalTime
            End Get
        End Property

        Public ReadOnly Property TimeOnly() As String
            Get
                Return notesDateTime.TimeOnly
            End Get
        End Property

        Public ReadOnly Property TimeZone() As Integer
            Get
                Return notesDateTime.TimeZone
            End Get
        End Property

        Public ReadOnly Property ZoneTime() As String
            Get
                Return notesDateTime.ZoneTime
            End Get
        End Property
        
    End Class
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Public Class NotesUIWorkspace

        Private notesUIWorkspace As Object
        
        Public Sub New()
            Me.notesUIWorkspace = CreateObject("Notes.NotesUIWorkspace")
        End Sub
        
        Protected Overrides Sub Finalize()
            Me.notesUIWorkspace = Nothing
        End Sub

        Public ReadOnly Property CurrentDatabase() As NotesUIDatabase
            Get
                Return New LN.NotesUIDatabase(Me.notesUIWorkspace.CurrentDatabase)
            End Get
        End Property

    End Class
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Public Class NotesUIDatabase

        Private notesUIDatabase As Object
        
        Public Sub New(ByRef notesUIDatabase As Object)
            Me.notesUIDatabase = notesUIDatabase
        End Sub
        
        Protected Overrides Sub Finalize()
            Me.notesUIDatabase = Nothing
        End Sub

        Public ReadOnly Property Database() As NotesDatabase
            Get
                Return notesUIDatabase.Database
            End Get
        End Property

        Public ReadOnly Property Documents() As NotesDocumentCollection
            Get
                Return New LN.NotesDocumentCollection(notesUIDatabase.Documents)
            End Get
        End Property

    End Class
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Public Class NotesUIView

        Private notesUIView As Object
        
        Public Sub New(ByRef notesUIView As Object)
            Me.notesUIView = notesUIView
        End Sub
        
        Protected Overrides Sub Finalize()
            Me.notesUIView = Nothing
        End Sub
        
    End Class
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Public Class NotesUIDocument

        Private notesUIDocument As Object
        
        Public Sub New(ByRef notesUIDocument As Object)
            Me.notesUIDocument = notesUIDocument
        End Sub
        
        Protected Overrides Sub Finalize()
            Me.notesUIDocument = Nothing
        End Sub
        
    End Class
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

End Namespace
