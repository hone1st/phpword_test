<?php

use PhpOffice\PhpWord\IOFactory;

/**
 * Created by PhpStorm.
 * User: ouxuan
 * Date: 2019/7/31
 * Time: 10:24
 */
class MyPhpWord
{
    public $zw_font_style = null;  // 正文字体
    public $zw_font_bold_style = null; // 正文加粗字体
    public $zw_para_style = null; // 正文段落样式
    public $bt_font_style = null; // 标题字体
    public $bt_para_style = null; // 标题段落
    private $phpWord = null;

    private $defaultTableStyle = [
        'borderColor' => '006699',
        'borderSize' => 6,  //
        'cellMargin' => 50,
        "alignment" => "center"
    ];

    /**
     * Left to Right, Top to Bottom
     */
    const TEXT_DIR_LRTB = 'lrTb';
    /**
     * Top to Bottom, Right to Left
     */
    const TEXT_DIR_TBRL = 'tbRl';
    /**
     * Bottom to Top, Left to Right
     */
    const TEXT_DIR_BTLR = 'btLr';
    /**
     * Left to Right, Top to Bottom Rotated
     */
    const TEXT_DIR_LRTBV = 'lrTbV';
    /**
     * Top to Bottom, Right to Left Rotated
     */
    const TEXT_DIR_TBRLV = 'tbRlV';
    /**
     * Top to Bottom, Left to Right Rotated
     */
    const TEXT_DIR_TBLRV = 'tbLrV';


    public function __construct(\PhpOffice\PhpWord\PhpWord $phpWord = null)
    {
        $phpWord = new \PhpOffice\PhpWord\PhpWord();
        if ($phpWord !== null) {
            $this->phpWord = $phpWord;
        }
    }

    public function getNewSection($style = null)
    {
        return $this->phpWord->addSection($style);
    }

    public function getPhpWord()
    {
        return $this->phpWord;
    }


    /**
     * 获取行的样式
     * @param bool $cantSplit 当行不够长的时候 是否可以下一行显示
     * @param int $exactHeigh 准确的行高
     * @param bool $tblHeader 设置该行是否在分页未展示完该表格的时候重复显示第一行
     * @return array
     */
    public function getRowStyle(bool $cantSplit = false, int $exactHeigh = 0, bool $tblHeader = false)
    {
        $row_style = [
            'cantSplit' => false,
            'tblHeader' => false
        ];

        if ($cantSplit === true) {
            $row_style['cantSplit'] = true;
        }
        if ($exactHeigh !== 0) {
            $row_style['exactHeigh'] = $exactHeigh;
        }

        if ($tblHeader === true) {
            $row_style['tblHeader'] = true;
        }

        return $row_style;
    }

    /**
     * 返回表格的样式
     * @param string $alignment 表格位置
     * @param string|null $bgColor 表格背景颜色
     * @param string $borderColor 表格边框颜色
     * @param float|null $borderSize 表格边框大小
     * @param float|null $cellMargin 列和列之间的间隔
     * @return array
     */
    public function getTableStyle(string $alignment = "center", string $bgColor = null, string $borderColor = null, float $borderSize = null, float $cellMargin = null)
    {
        $table_style = [];
        if (in_array($alignment, ["left", "center", "start", "end", "both", "mediumKashida", "distribute", "numTab", "highKashida", "lowKashida", "thaiDistribute", "right", "justify"])) {
            $p_style['alignment'] = $alignment;
        }
        if ($bgColor !== null) {
            $table_style['bgColor'] = $bgColor;
        }
        if ($borderColor !== null) {
            $table_style['bgColor'] = $borderColor;
        }
        if ($borderSize !== null) {
            $table_style['bgColor'] = $this->liMiToTwip($cellMargin);
        }
        if ($cellMargin !== null) {
            $table_style['bgColor'] = $this->liMiToTwip($cellMargin);
        }

        return $table_style;

    }

    /**
     * 获取列的样式
     * @param string $bgColor 背景颜色
     * @param int $gridSpan
     * @param string|null $borderColor 边框颜色
     * @param int $borderSize 边框大小
     * @param string $textDirection 文本展示的方向
     * @param string $valign 垂直方向   top|center|both|bottom
     * @param int $width 列的大小
     * @param string $vMerge 和平还是重新开始    restart|continue
     * @return array
     */
    public function getCellStyle(string $bgColor = "006699", string $valign = null, int $gridSpan = 0, string $borderColor = null, int $borderSize = 0, string $textDirection = null, int $width = 0, string $vMerge = null): array
    {
        // 默认的列样式
        $table_cell_style = [
            'bgColor' => $bgColor
        ];
        if (in_array($textDirection, [self::TEXT_DIR_BTLR, self::TEXT_DIR_LRTB, self::TEXT_DIR_TBLRV, self::TEXT_DIR_TBRLV, self::TEXT_DIR_TBRL, self::TEXT_DIR_LRTBV])) {
            $table_cell_style['textDirection'] = $textDirection;
        }
        if (in_array($valign, ["top", "center", "both", "bottom"])) {
            $table_cell_style['valign'] = $valign;
        }
        if (in_array($vMerge, ["restart", "continue"])) {
            $table_cell_style['vMerge'] = $vMerge;
        }
        if ($gridSpan !== 0) {
            $table_cell_style['gridSpan'] = $gridSpan;
        }
        if ($borderColor !== null) {
            $table_cell_style['borderColor'] = $borderColor;
        }
        if ($borderSize !== 0) {
            $table_cell_style['borderSize'] = $borderSize;
        }

        return $table_cell_style;
    }

    /**
     * 获取字体样式
     * @param string $name 字体名字
     * @param string $size 字体大小
     * @param bool $bold 是否加粗
     * @param string $color 字体颜色
     * @param bool $superScript 上标
     * @return array
     */
    public function getFontStyle(string $name = null, string $size = null, bool $bold = false, string $color = null, bool $superScript = false)
    {
        // 默认字体样式
        $font_style = [
            "name" => "宋体",
            "size" => $this->getFontSize("小四")
        ];

        if ($name !== null) {
            $font_style['name'] = $name;
        }
        if ($size !== null) {
            $font_style['size'] = $this->getFontSize($size);
        }
        if ($bold === true) {
            $font_style['bold'] = $bold;
        }
        if ($superScript === true) {
            $font_style['superScript'] = $superScript;
        }
        if ($color !== null) {
            $font_style['color'] = $color;
        }
//        \PhpOffice\PhpWord\SimpleType\TextAlignment::
        return $font_style;
    }

    /**
     * @param string|null $alignment
     * @param float|null $height
     * @param float|null $width
     * @param string|null $wrappingStyle
     * @param string|null $positioning
     * @param float|null $marginTop
     * @param float|null $marginLeft
     * @return array
     */
    public function getImageStyle(string $alignment = null, float $height = null, float $width = null, string $wrappingStyle = null, string $positioning = null, float $marginTop = null, float $marginLeft = null)
    {
        $imageStyle = [];
        if (in_array($alignment, ["left", "center", "start", "end", "both", "mediumKashida", "distribute", "numTab", "highKashida", "lowKashida", "thaiDistribute", "right", "justify"])) {
            $imageStyle['alignment'] = $alignment;
        }
        if ($height !== null) {
            $imageStyle['height'] = $height;
        }
        if (in_array($wrappingStyle, ["inline", "square", "tight", "behind", "infront"])) {
            $imageStyle['wrappingStyle'] = $wrappingStyle;
        }
        if ($positioning !== null) {
            $imageStyle['positioning'] = $positioning;
        }
        if ($width !== null) {
            $imageStyle['width'] = $width;
        }

        if ($marginLeft !== null) {
            $imageStyle['marginLeft'] = $marginLeft;
        }
        if ($marginTop !== null) {
            $imageStyle['marginTop'] = $marginTop;
        }

        return $imageStyle;
    }

    /**
     * 缩进字符 字号转twip
     * @param string $fontSize
     * @return float
     */
    public function firstIndent(string $fontSize)
    {
        $fontZF = [
            "小五" => 0.23,
            "五号" => 0.27,
            "小四" => 0.31,
            "四号" => 0.35,
            "小三" => 0.37,
            "三号" => 0.40,
            "小二" => 0.45,
            "二号" => 0.55,
            "小一" => 0.60,
            "一号" => 0.65,
            "小初" => 0.85,
            "初号" => 1.00,
        ];
        return $this->liMiToTwip($fontZF[$fontSize]);
    }

    /**
     * 返回文本样式段落
     * @param string $align
     * @param float|null $spacing $spacingLineRule 不设置就是多倍 exact则此值为固定值
     * @param array|null $indentation ["firstLine"=>,"left"=>,"right"=>,"hanging"=>]   首行 左 右 悬挂
     * @param bool $keepLines
     * @param bool $keepNext
     * @param float|null $lineHeight 多倍行距
     * @param float|null $spaceAfter
     * @param float|null $spaceBefore
     * @param bool $widowControl
     * @param string $spacingLineRule
     * @param bool $next 是否适用于下一段落的样式
     * @return array
     */
    public function getParagraphStyle(string $align = "center", float $spacing = null, array $indentation = null, bool $keepLines = false, bool $keepNext = false, float $lineHeight = null, float $spaceAfter = null, float $spaceBefore = null, bool $widowControl = false, string $spacingLineRule = null, bool $next = false)
    {
        $p_style = [
            'align' => "center"
        ];

        if (in_array($align, ["left", "center", "start", "end", "both", "mediumKashida", "distribute", "numTab", "highKashida", "lowKashida", "thaiDistribute", "right", "justify"])) {
            $p_style['align'] = $align;
        }
        if ($keepLines === true) {
            $p_style['keepLines'] = $keepLines;
        }
        if ($next === true) {
            $p_style['next'] = $next;
        }
        if ($keepNext === true) {
            $p_style['keepNext'] = $keepNext;
        }
        if ($indentation !== null) {
            $p_style['indentation'] = $indentation;
        }
        if ($widowControl === true) {
            $p_style['widowControl'] = $widowControl;
        }
        if (in_array($spacingLineRule, ["auto", "exact", "atLeast"])) {
            $p_style['spacingLineRule'] = $spacingLineRule;
        }

        if ($lineHeight !== null) {
            $p_style['lineHeight'] = $lineHeight;
        }

        if ($spacing !== null) {
            $p_style['spacing'] = $this->liMiToTwip($spacing);
        }
        if ($spaceAfter !== null) {
            $p_style['spaceAfter'] = $this->liMiToTwip($spaceAfter / 1.073 * 0.035);
        }
        if ($spaceBefore !== null) {
            $p_style['spaceBefore'] = $this->liMiToTwip($spaceBefore / 1.073 * 0.035);
        }
        return $p_style;
    }

    /**
     * 厘米换成twip
     * @param float $limi
     * @return float
     */
    public function liMiToTwip(float $limi): float
    {
        return 567 * $limi;
    }

    /**
     * 返回文字映射大小
     * @param string $size
     * @return mixed|string
     */
    public function getFontSize(string $size)
    {
        $font_size_str_to_number = [
            "初号" => 44,
            "小初" => 36,
            "一号" => 26,
            "小一" => 24,
            "二号" => 22,
            "小二" => 18,
            "三号" => 16,
            "小三" => 15,
            "四号" => 14,
            "小四" => 12,
            "五号" => 10.5,
            "小五" => 9,
            "六号" => 7.5,
            "小六" => 6.5,
            "七号" => 5.5,
            "八号" => 5
        ];

        return empty($font_size_str_to_number[$size]) ? $size : $font_size_str_to_number[$size];

    }

    /**
     * 保存
     * @param string $path
     * @throws \PhpOffice\PhpWord\Exception\Exception
     */
    public function save(string $path = "")
    {
        if ($path === "") {
            $name = date("Y-m-d-His");
            $path = "./" . $name . ".doc";
        }
        $writer = IOFactory::createWriter($this->phpWord, 'Word2007');
        $writer->save($path);
    }

    /**
     * 下载
     * @throws \PhpOffice\PhpWord\Exception\Exception
     */
    public function outPut()
    {
        $writer = IOFactory::createWriter($this->phpWord, 'Word2007');
        $name = date("Y-m-d-His");
        header("Content-Description: File Transfer");
        header('Content-Disposition: attachment; filename="' . $name . ".docx" . '"');
        header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        header('Content-Transfer-Encoding: binary');
        header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
        header('Expires: 0');
        $writer->save("php://output");
    }

    /**
     * @param array $col_cn_field 姓名 => name, 性别 => sex , 年龄 => age
     * @param array $col_field_data ['name'=>"张三",sex=> 男, age = 23]
     * @param int|array $cellWidth
     * @param array|null $col_style
     * @param array|null $content_style
     * @param array|null $p_style
     * @param \PhpOffice\PhpWord\Element\Section|null $section
     * @param array|null $tableStyle
     * @return \PhpOffice\PhpWord\Element\Table
     */
    public function tableColHeader(array $col_cn_field, array $col_field_data, array $cellWidth = null, array $col_style = null, array $content_style = null, array $p_style = null, \PhpOffice\PhpWord\Element\Section $section = null, array $tableStyle = null)
    {
        if ($section === null) {
            $section = $this->phpWord->addSection();
        }
        if ($tableStyle === null) {
            $tableStyle = $this->defaultTableStyle;
        }
        $table = $section->addTable($tableStyle);
        foreach ($col_cn_field as $cn => $field) {
            $row = $table->addRow($this->liMiToTwip(0.5));
            $row->addCell($this->liMiToTwip($cellWidth[$cn]), $col_style[$cn])->addText($cn, $content_style[$cn], $p_style[$cn]);
            if (ord(str_split($field, 1)[0]) >= 60 && ord(str_split($field, 1)[0]) <= 90) {
                try {
                    // 匹配该字段的的值是否是张有效的图片
//                   $check = preg_match_all('/\.jpg$/', $col_field_data[lcfirst($field)]);
                    $this->checkImage($col_field_data[lcfirst($field)]);
                    $row->addCell($this->liMiToTwip($cellWidth[$field]), $col_style[$field])->addImage($col_field_data[lcfirst($field)], $content_style[$field]);
                } catch (Exception $e) {
                    $row->addCell($this->liMiToTwip($cellWidth[$field]), $col_style[$field])->addImage("http://dd.falv58.com/Common/upload/cos/2019-07-1215473864691641.png", $content_style[$field]);
                }
            } else {
                $row->addCell($this->liMiToTwip($cellWidth[$field]), $col_style[$field])->addText($col_field_data[$field], $content_style[$field], $p_style[$field]);
            }
        }
        return $table;
    }

    /**
     * 检测图片是否存在
     * @param $source
     * @throws \PhpOffice\PhpWord\Exception\InvalidImageException
     * @throws \PhpOffice\PhpWord\Exception\UnsupportedImageTypeException
     */
    private function checkImage($source)
    {
        new \PhpOffice\PhpWord\Element\Image($source);
    }

    /**
     * @param float $headerHeight
     * @param array $row_cn
     * @param array $cellWidth
     * @param array $row_datas
     * @param array|null $fields
     * @param array|null $cell_style
     * @param array|null $content_style
     * @param array|null $p_style
     * @param \PhpOffice\PhpWord\Element\Section|null $section
     * @param array|null $tableStyle
     */
    public function tableRowHeader(float $headerHeight, array $row_cn, array $cellWidth, array $row_datas, array $fields = null, array $cell_style = null, array $content_style = null, array $p_style = null, \PhpOffice\PhpWord\Element\Section $section = null, array $tableStyle = null)
    {
        if ($section === null) {
            $section = $this->phpWord->addSection();
        }
        if ($tableStyle === null) {
            $tableStyle = $this->defaultTableStyle;
        }
        $table = $section->addTable($tableStyle);
        // tblHeader 跨页重新展示的表头
        $row = $table->addRow($this->liMiToTwip($headerHeight), ['tblHeader' => true]);
        foreach ($row_cn as $cn) {
            $row->addCell($this->liMiToTwip($cellWidth[$cn]), $cell_style[$cn])->addText($cn, $content_style[$cn], $p_style[$cn]);
        }
        foreach ($row_datas as $item) {
            $row = $table->addRow($this->liMiToTwip($headerHeight));
            foreach ($fields as $field) {
                if (ord(str_split($field, 1)[0]) >= 60 && ord(str_split($field, 1)[0]) <= 90) {
                    try {
//                        $check = preg_match_all('/\.png/', $item[$field]);
                        // 匹配该字段的的值是否是张有效的图片
                        $this->checkImage($item[$field]);
                        $row->addCell($this->liMiToTwip($cellWidth[$field]))->addImage($item[$field], $content_style[$field]);
                    } catch (Exception $e) {
                        $row->addCell($this->liMiToTwip($cellWidth[$field]))->addImage("http://dd.falv58.com/Common/upload/cos/2019-07-1215473864691641.png", $content_style[$field]);
                    }
                } else {
                    $row->addCell($this->liMiToTwip($cellWidth[$field]), $cell_style[$field])
                        ->addText($item[$field], $content_style[$field], $p_style[$field]);
                }
            }
        }
    }

    public function __destruct()
    {
        $this->phpWord = null;
    }
}