from pydantic import BaseModel, Field, constr, model_validator
from typing import List, Optional, Literal, Union, Any

# 型定義
HorizontalAlign = Literal["left", "center", "right"]
VerticalAlign = Literal["top", "middle", "bottom"]
ColorHex = constr(pattern=r"^#([0-9A-Fa-f]{6})$")  # "#RRGGBB"

class CellStyle(BaseModel):
    fill: Optional[ColorHex] = Field(default=None, description="背景色")
    text_color: Optional[ColorHex] = Field(default=None, description="文字色")
    bold: Optional[bool] = Field(default=None, description="太字")
    italic: Optional[bool] = Field(default=None, description="斜体")
    underline: Optional[bool] = Field(default=None, description="下線")
    align: Optional[HorizontalAlign] = Field(default=None, description="水平寄せ（セル単位）")
    vertical_align: Optional[VerticalAlign] = Field(default=None, description="縦寄せ（セル単位）")
    font_size: Optional[int] = Field(default=None, gt=0, description="フォントサイズ（pt）")

    model_config = {"from_attributes": True, "extra": "allow"}

class TableCell(BaseModel):
    value: Any = Field(default="")
    # ショートハンド: {"value": "...", "fill": "#..."} も許可
    fill: Optional[ColorHex] = None
    text_color: Optional[ColorHex] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[bool] = None
    align: Optional[HorizontalAlign] = None
    vertical_align: Optional[VerticalAlign] = None
    font_size: Optional[int] = Field(default=None, gt=0)
    style: Optional[CellStyle] = Field(default=None, description="詳細スタイル（fill等と併用可）")

    model_config = {"from_attributes": True, "extra": "allow"}

CellScalar = Union[str, int, float, bool]
CellLike = Union[CellScalar, TableCell]

class TableSchema(BaseModel):
    """
    table.py レンダラ用データスキーマ（セル単位で色/スタイルを直指定できる版）
    """

    # ---- データ本体（セルはスカラー or TableCell）----
    column_headers: Optional[List[CellLike]] = Field(default=None, description="上ヘッダー（列名）")
    row_headers: Optional[List[CellLike]] = Field(default=None, description="左ヘッダー（行名）")
    rows: List[List[CellLike]] = Field(default_factory=list, description="セルの2次元配列（行ベース）")

    # ---- 行ゼブラ（任意）----
    banding: bool = Field(default=True, description="交互帯（行方向固定）を有効化")
    band_start_index: int = Field(default=0, ge=0, description="交互開始のオフセット（データ行 0 起点）")

    # ---- 既定の配色/スタイル（セル個別の指定があればそちらが優先）----
    header_fill: Optional[ColorHex] = Field(default="#DDEBF7", description="ヘッダー背景色")
    header_text_color: Optional[ColorHex] = Field(default="#0F172A", description="ヘッダーテキスト色")
    band_fill: Optional[ColorHex] = Field(default="#F2F2F2", description="交互帯の塗り色")
    cell_text_color: Optional[ColorHex] = Field(default="#111111", description="本文の既定文字色")
    table_style: Optional[str] = Field(default=None, description='python-pptx のビルトイン表スタイル（例: "Table Grid"）')

    # ---- テキスト／レイアウト（既定値）----
    font_size: int = Field(default=11, gt=0, description="本文フォントサイズ（pt）")
    header_font_size: int = Field(default=12, gt=0, description="ヘッダーフォントサイズ（pt）")
    align: HorizontalAlign = Field(default="left", description="本文の既定水平寄せ")
    col_align: Optional[List[HorizontalAlign]] = Field(
        default=None, description="列ごとの水平寄せ（行ヘッダー列を含む総列数に合わせる）"
    )
    wrap: bool = Field(default=True, description="テキストの折返し")
    vertical_align: VerticalAlign = Field(default="middle", description="セルの既定縦寄せ")
    cell_padding_pt: int = Field(default=4, ge=0, description="セルの余白（pt）")

    # ---- 寸法（重み配分）----
    col_widths: Optional[List[float]] = Field(default=None, description="列幅の重み（長さ=総列数）")
    row_heights: Optional[List[float]] = Field(default=None, description="行高の重み（長さ=総行数）")

    model_config = {"from_attributes": True, "extra": "allow"}

    @model_validator(mode="after")
    def _sanitize(self):
        if self.band_start_index is None or self.band_start_index < 0:
            self.band_start_index = 0
        return self

# Export 統一名
Schema = TableSchema
