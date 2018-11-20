Imports System
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports Microsoft.VisualBasic.Interaction

Namespace LN

    Public Enum ACLLEVEL As Integer
        ACLLEVEL_AUTHOR = 3
        ACLLEVEL_DEPOSITOR = 1
        ACLLEVEL_DESIGNER = 5
        ACLLEVEL_EDITOR = 4
        ACLLEVEL_MANAGER = 6
        ACLLEVEL_NOACCESS = 0
        ACLLEVEL_READER = 2
    End Enum

    Public Enum ACLTYPE As Integer
        ACLTYPE_MIXED_GROUP = 3
        ACLTYPE_PERSON = 1
        ACLTYPE_PERSON_GROUP = 4
        ACLTYPE_SERVER = 2
        ACLTYPE_SERVER_GROUP = 5
        ACLTYPE_UNSPECIFIED = 0
    End Enum

    Public Enum AG_TARGET As Integer
        TARGET_ALL_DOCS = 1
        TARGET_ALL_DOCS_IN_VIEW = 5
        TARGET_NEW_DOCS = 2
        TARGET_NEW_OR_MODIFIED_DOCS = 3
        TARGET_NONE = 0
        TARGET_SELECTED_DOCS = 4
        TARGET_UNREAD_DOCS_IN_VIEW = 6
    End Enum

    Public Enum AG_TRIGGER As Integer
        TRIGGER_AFTER_MAIL_DELIVERY = 2
        TRIGGER_BEFORE_MAIL_DELIVERY = 6
        TRIGGER_DOC_PASTED = 3
        TRIGGER_MANUAL = 4
        TRIGGER_NONE = 0
        TRIGGER_SCHEDULED = 1
        TRIGGER_UPDATE = 5
    End Enum

    Public Enum COLORS As Integer
        COLOR_BLACK = 0
        COLOR_BLUE = 4
        COLOR_CYAN = 7
        COLOR_DARK_BLUE = 10
        COLOR_DARK_CYAN = 13
        COLOR_DARK_GREEN = 9
        COLOR_DARK_MAGENTA = 11
        COLOR_DARK_RED = 8
        COLOR_DARK_YELLOW = 12
        COLOR_GRAY = 14
        COLOR_GREEN = 3
        COLOR_LIGHT_GRAY = 15
        COLOR_MAGENTA = 5
        COLOR_RED = 2
        COLOR_WHITE = 1
        COLOR_YELLOW = 6
    End Enum

    Public Enum DB_TYPES As Integer
        DATABASE = 1247 ' avoid Using this In COM
        NOTES_DATABASE = 1247
        REPLICA_CANDIDATE = 1245
        TEMPLATE = 1248
        TEMPLATE_CANDIDATE = 1246
    End Enum

    Public Enum EMBED_TYPE As Integer
        EMBED_ATTACHMENT = 1454
        EMBED_OBJECT = 1453
        EMBED_OBJECTLINK = 1452
    End Enum

    Public Enum FT_TYPES As Integer
        FT_DATABASE = 8192
        FT_DATE_ASC = 64
        FT_DATE_DES = 32
        FT_FILESYSTEM = 4096
        FT_FUZZY = 16384
        FT_SCORES = 8
        FT_STEMS = 512
        FT_THESAURUS = 1024
    End Enum

    Public Enum IT_TYPE As Integer
        ACTIONCD = 16
        ASSISTANTINFO = 17
        ATTACHMENT = 1084
        AUTHORS = 1076
        COLLATION = 2
        DATETIMES = 1024
        EMBEDDEDOBJECT = 1090
        ERRORITEM = 256
        FORMULA = 1536
        HTML = 21
        ICON = 6
        LSOBJECT = 20
        MIME_PART = 25
        NAMES = 1074
        NOTELINKS = 7
        NOTEREFS = 4
        NUMBERS = 768
        OTHEROBJECT = 1085
        QUERYCD = 15
        READERS = 1075
        RICHTEXT = 1
        SIGNATURE = 8
        TEXT = 1280
        UNAVAILABLE = 512
        UNKNOWN = 0
        USERDATA = 14
        USERID = 1792
        VIEWMAPDATA = 18
        VIEWMAPLAYOUT = 19
    End Enum

    Public Enum LOG_EVENTS As Integer
        EV_ALARM = 8
        EV_COMM = 1
        EV_MAIL = 3
        EV_MISC = 6
        EV_REPLICA = 4
        EV_RESOURCE = 5
        EV_SECURITY = 2
        EV_SERVER = 7
        EV_UNKNOWN = 0
        EV_UPDATE = 9
    End Enum

    Public Enum LOG_SEVERITY As Integer
        SEV_FAILURE = 2
        SEV_FATAL = 1
        SEV_NORMAL = 5
        SEV_UNKNOWN = 0
        SEV_WARNING1 = 3
        SEV_WARNING2 = 4
    End Enum

    Public Enum OE_CLASS As Integer
        OUTLINE_CLASS_DATABASE = 2194
        OUTLINE_CLASS_DOCUMENT = 2190
        OUTLINE_CLASS_FOLDER = 2197
        OUTLINE_CLASS_FORM = 2192
        OUTLINE_CLASS_FRAMESET = 2195
        OUTLINE_CLASS_NAVIGATOR = 2193
        OUTLINE_CLASS_PAGE = 2196
        OUTLINE_CLASS_UNKNOWN = 2189
        OUTLINE_CLASS_VIEW = 2191
    End Enum

    Public Enum OE_TYPE As Integer
        OUTLINE_OTHER_FOLDERS_TYPE = 1589
        OUTLINE_OTHER_UNKNOWN_TYPE = 1591
        OUTLINE_OTHER_VIEWS_TYPE = 1588
        OUTLINE_TYPE_ACTION = 2188
        OUTLINE_TYPE_NAMEDELEMENT = 2187
        OUTLINE_TYPE_NOTELINK = 2186
        OUTLINE_TYPE_URL = 2185
    End Enum

    Public Enum REG_TYPE As Integer
        ID_CERTIFIER = 173
        ID_FLAT = 171
        ID_HIERARCHICAL = 172
    End Enum

    Public Enum RP_PRIORITY As Integer
        DB_REPLICATION_PRIORITY_HIGH = 1549
        DB_REPLICATION_PRIORITY_LOW = 1547
        DB_REPLICATION_PRIORITY_MED = 1548
        DB_REPLICATION_PRIORITY_NOTSET = 1565
    End Enum

    Public Enum RT_ALIGN As Integer
        ALIGN_CENTER = 3
        ALIGN_FULL = 2
        ALIGN_LEFT = 0
        ALIGN_NOWRAP = 4
        ALIGN_RIGHT = 1
    End Enum

    Public Enum RT_EFFECTS As Integer
        EFFECTS_EMBOSS = 4
        EFFECTS_EXTRUDE = 5
        EFFECTS_NONE = 0
        EFFECTS_SHADOW = 3
        EFFECTS_SUBSCRIPT = 2
        EFFECTS_SUPERSCRIPT = 1
    End Enum

    Public Enum RT_FONTS As Integer
        FONT_COURIER = 4
        FONT_HELV = 1
        FONT_ROMAN = 0
    End Enum

    Public Enum RT_PAGINATE As Integer
        PAGINATE_BEFORE = 1
        PAGINATE_DEFAULT = 0
        PAGINATE_KEEP_TOGETHER = 4
        PAGINATE_KEEP_WITH_NEXT = 2
    End Enum

    Public Enum RT_TAB As Integer
        TAB_CENTER = 3
        TAB_DECIMAL = 2
        TAB_LEFT = 0
        TAB_RIGHT = 1
    End Enum

    Public Enum SPACING As Integer
        SPACING_DOUBLE = 4
        SPACING_ONE_POINT_25 = 1
        SPACING_ONE_POINT_50 = 2
        SPACING_ONE_POINT_75 = 3
        SPACING_SINGLE = 0
    End Enum

    Public Enum USER_TYPE As Integer
        NOTES_DESKTOP_CLIENT = 175
        NOTES_FULL_CLIENT = 176
        NOTES_LIMITED_CLIENT = 174
    End Enum

    Public Enum VC_ALIGN As Integer
        VC_ALIGN_CENTER = 2
        VC_ALIGN_LEFT = 0
        VC_ALIGN_RIGHT = 1
    End Enum

    Public Enum VC_DATEFMT As Integer
        VC_FMT_MD = 2
        VC_FMT_Y4M = 6
        VC_FMT_YM = 3
        VC_FMT_YMD = 0
    End Enum

    Public Enum VC_FONTSTYLE As Integer
        VC_FONT_BOLD = 1
        VC_FONT_ITALIC = 2
        VC_FONT_STRIKEOUT = 8
        VC_FONT_UNDERLINE = 4
    End Enum

    Public Enum VC_NUMATTR As Integer
        VC_ATTR_PARENS = 2
        VC_ATTR_PERCENT = 4
        VC_ATTR_PUNCTUATED = 1
    End Enum

    Public Enum VC_NUMFMT As Integer
        VC_FMT_CURRENCY = 3
        VC_FMT_FIXED = 1
        VC_FMT_GENERAL = 0
        VC_FMT_SCIENTIFIC = 2
    End Enum

    Public Enum VC_SEP As Integer
        VC_SEP_COMMA = 2
        VC_SEP_NEWLINE = 4
        VC_SEP_SEMICOLON = 3
        VC_SEP_SPACE = 1
    End Enum

    Public Enum VC_TDFMT As Integer
        VC_FMT_DATE = 0
        VC_FMT_DATETIME = 2
        VC_FMT_TIME = 1
        VC_FMT_TODAYTIME = 3
    End Enum

    Public Enum VC_TIMEFMT As Integer
        VC_FMT_HM = 1
        VC_FMT_HMS = 0
    End Enum

    Public Enum VC_TIMEZONEFMT As Integer
        VC_FMT_ALWAYS = 2
        VC_FMT_NEVER = 0
        VC_FMT_SOMETIMES = 1
    End Enum

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
        TEXT = 1281
        FORMULA = 1536
        USERID = 1792
    End Enum

End Namespace
