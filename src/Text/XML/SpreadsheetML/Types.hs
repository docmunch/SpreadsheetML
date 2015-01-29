{-# LANGUAGE GeneralizedNewtypeDeriving #-}
module Text.XML.SpreadsheetML.Types where

{- See http://msdn.microsoft.com/en-us/library/aa140066%28office.10%29.aspx -}

import Data.Word ( Word64 )
import Data.Time.Clock (UTCTime)
import Data.Time.Format (formatTime, readsTime, ParseTime, FormatTime)
import System.Locale (defaultTimeLocale, iso8601DateFormat)

-- | Only implement what we need

data Workbook = Workbook
  { workbookDocumentProperties :: Maybe DocumentProperties
  , workbookStyles             :: Maybe Styles
  , workbookWorksheets         :: [Worksheet]
  }
  deriving (Read, Show)

data Styles = Styles
  { elemStyles :: [Style]
  }
  deriving (Read, Show)
  
data Style = Style
  -- attributes
  { attribID         :: StyleID
  , attribName       :: Maybe String
  , attribParent     :: Maybe String
  -- elements
  , elemAlignment    :: Maybe Alignment
  , elemBorders      :: Maybe Borders
  , elemFont         :: Maybe Font
  , elemInterior     :: Maybe Interior
  , elemNumberFormat :: Maybe NumberFormat
  , elemProtection   :: Maybe Protection
  }
  deriving (Read, Show)
  
data Alignment = Alignment
  { alignHorizontal   :: Maybe Horizontal
  , alignReadingOrder :: Maybe ReadingOrder
  , alignRotate       :: Maybe Double
  , alignShrinkToFit  :: Maybe Bool
  , alignVertical     :: Maybe Vertical
  }
  deriving (Read, Show)
  
data Borders = Borders
  { elemBorder :: [Border] }
  deriving (Read, Show)
  
data Border = Border
  { attribPosition    :: Maybe Position
  , attribBorderColor :: Maybe String
  , attribLineStyle   :: Maybe LineStyle
  , attribWeight      :: Maybe LineWeight
  }
  deriving (Read, Show)
  
data Font = Font
  { attribBold          :: Maybe Bool
  , attribFontColor     :: Maybe String
  , attribFontName      :: Maybe String
  , attribItalic        :: Maybe Bool
  , attribSize          :: Maybe Double
  , attribStrikeThrough :: Maybe Bool
  , attribUnderline     :: Maybe Underline
  , attribCharSet       :: Maybe Word64
  , attribFamily        :: Maybe FontFamily
  }
  deriving (Read, Show)
  
data Interior = Interior
  { attribInteriorColor :: Maybe String
  , attribPattern       :: Maybe Pattern
  }
  deriving (Read, Show)
  
data NumberFormat = NumberFormat
  { attribFormat :: Maybe String }
  deriving (Read, Show)
  
data Protection = Protection
  { attribProtected   :: Maybe Bool
  , attribHideFormula :: Maybe Bool
  }
  deriving (Read, Show)
  
data DocumentProperties = DocumentProperties
  { documentPropertiesTitle       :: Maybe String
  , documentPropertiesSubject     :: Maybe String
  , documentPropertiesKeywords    :: Maybe String
  , documentPropertiesDescription :: Maybe String
  , documentPropertiesRevision    :: Maybe Word64
  , documentPropertiesAppName     :: Maybe String
  , documentPropertiesCreated     :: Maybe String -- ^ Actually, this should be a date time
  }
  deriving (Read, Show)

data Worksheet = Worksheet
  { worksheetTable       :: Maybe Table
  , worksheetName        :: Name
  }
  deriving (Read, Show)

data Table = Table
  { tableColumns             :: [Column]
  , tableRows                :: [Row]
  , tableStyleID             :: Maybe StyleID -- ^ Must be defined in Styles
  , tableDefaultColumnWidth  :: Maybe Double -- ^ Default is 48
  , tableDefaultRowHeight    :: Maybe Double -- ^ Default is 12.75
  , tableExpandedColumnCount :: Maybe Word64
  , tableExpandedRowCount    :: Maybe Word64
  , tableLeftCell            :: Maybe Word64 -- ^ Default is 1
  , tableTopCell             :: Maybe Word64 -- ^ Default is 1
  , tableFullColumns         :: Maybe Bool
  , tableFullRows            :: Maybe Bool
  }
  deriving (Read, Show)

data Column = Column
  { columnCaption      :: Maybe Caption
  , columnStyleID      :: Maybe StyleID -- ^ Must be defined in Styles
  , columnAutoFitWidth :: Maybe AutoFitWidth
  , columnHidden       :: Maybe Hidden
  , columnIndex        :: Maybe Word64
  , columnSpan         :: Maybe Word64
  , columnWidth        :: Maybe Double
  }
  deriving (Read, Show)

data Row = Row
  { rowCells         :: [Cell]
  , rowStyleID       :: Maybe StyleID -- ^ Must be defined in Styles
  , rowCaption       :: Maybe Caption
  , rowAutoFitHeight :: Maybe AutoFitHeight
  , rowHeight        :: Maybe Double
  , rowHidden        :: Maybe Hidden
  , rowIndex         :: Maybe Word64
  , rowSpan          :: Maybe Word64
  }
  deriving (Read, Show)

data Cell = Cell
  -- elements
  { cellData          :: Maybe ExcelValue
  -- Attributes
  , cellStyleID       :: Maybe StyleID -- ^ Must be defined in Styles
  , cellFormula       :: Maybe Formula
  , cellIndex         :: Maybe Word64
  , cellMergeAcross   :: Maybe Word64
  , cellMergeDown     :: Maybe Word64
  }
  deriving (Read, Show)

data ExcelValue
    = Number Double
    | Boolean Bool
    | StringType String
    | UtcType WrappedUtc -- ^ be sure to set the other attributes: cellStyleId = "s64", cellIndex = 7 }
  deriving (Read, Show)

-- Necessary since UtcTime doens't have read/show instances
newtype WrappedUtc = WrappedUtc { toUtc :: UTCTime }
    deriving (FormatTime, ParseTime)
instance Show WrappedUtc where
    show s = formatTime defaultTimeLocale (iso8601DateFormat $ Just "%H:%M:%S") s
instance Read WrappedUtc where
    readsPrec _ =
        readsTime defaultTimeLocale (iso8601DateFormat $ Just "%H:%M:%S")

-- | TODO: Currently just a string, but we could model excel formulas and
-- use that type here instead.
newtype Formula = Formula String
  deriving (Read, Show)

data AutoFitWidth = AutoFitWidth | DoNotAutoFitWidth
  deriving (Read, Show)

data AutoFitHeight = AutoFitHeight | DoNotAutoFitHeight
  deriving (Read, Show)

-- | Attribute for hidden things
data Hidden = Shown | Hidden
  deriving (Read, Show)

data Horizontal = HAlignAutomatic | HAlignLeft | HAlignCenter | HAlignRight
  deriving (Read, Show)

data ReadingOrder = RightToLeft | LeftToRight
  deriving (Read, Show)
  
data Vertical = VAlignAutomatic | VAlignTop | VAlignBottom | VAlignCenter
  deriving (Read, Show)
  
data Position = PositionLeft | PositionTop | PositionRight | PositionBottom
  deriving (Read, Show)

data LineStyle = LineStyleNone | LineStyleContinuous | LineStyleDash | LineStyleDot | LineStyleDashDot | LineStyleDashDotDot
  deriving (Read, Show)

data LineWeight = Hairline | Thin | Medium | Thick
  deriving (Read, Show)

data Underline = UnderlineNone | UnderlineSingle | UnderlineDouble | UnderlineSingleAccounting | UnderlineDoubleAccounting
  deriving (Read, Show)

data FontFamily = Automatic | Decorative | Modern | Roman | Script | Swiss
  deriving (Read, Show)

data Pattern = PatternNone | PatternSolid | PatternGray75 | PatternGray50 | PatternGray25 | PatternGray125 | PatternGray0625 | 
               PatternHorzStripe | PatternVertStripe | PatternReverseDiagStripe | PatternDiagStripe | PatternDiagCross | 
               PatternThickDiagCross | PatternThinHorzStripe | PatternThinVertStripe | PatternThinReverseDiagStripe | 
               PatternThinDiagStripe | PatternThinHorzCross | PatternThinDiagCross
  deriving (Read, Show)
  
-- | For now this is just a string, but we could model excel's names
newtype Name = Name String
  deriving (Read, Show)

newtype Caption = Caption String
  deriving (Read, Show)

newtype StyleID = StyleID String
  deriving (Read, Show)

