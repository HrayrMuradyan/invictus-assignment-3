from docx.shared import Pt, Cm

class StyleConfig:
    """Centralized configuration for style rules."""

    FONT_NAME = "Arial"
    FONT_SIZE = Pt(9)
    
    # Cover Page Settings
    COVER_START_ROW = 19  
    COVER_TITLE_SIZE = Pt(14) 
    
    # Table Settings
    TABLE_ROW_HEIGHT = Cm(0.37)  
    TABLE_COLUMN_WIDTHS = [Cm(11.99), Cm(1.20), Cm(2.30), Cm(2.30)]
    TABLE_HANGING_INDENT = Cm(0.63)
    TABLE_CELL_MARGIN_SIDE = "100"  
    TABLE_CELL_MARGIN_TB = "0"  
