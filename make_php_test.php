<?php
/**
 * Created by PhpStorm.
 * User: STAM.LIANG
 * Date: 2019/7/1
 * Time: 15:08
 */

require "vendor/autoload.php";
require_once("./MyPhpWord.php");

//\PhpOffice\PhpWord\SimpleType\Jc::
$word = new MyPhpWord();
/**
 * 设置文档基础信息
 * @param MyPhpWord $word
 */
function setDocInfo(MyPhpWord $word)
{
    $properties = $word->getPhpWord()->getDocInfo();
    $properties->setCreator('STAM.LIANG');
    $properties->setCompany('XXXXX有限公司');
    $properties->setTitle('tile');
    $properties->setDescription('desc');
    $properties->setCategory('cate');
    $properties->setLastModifiedBy('STAM.LIANG');
    $properties->setCreated();
    $properties->setModified();
    $properties->setSubject('sub');
    $properties->setKeywords('key,word');
}

/**
 * 初始化样式
 * @param MyPhpWord $word
 */
function initStyleSetting(MyPhpWord $word)
{
    $fontName = "宋体";
    $fontSize = "小四";
    $word->getPhpWord()->getSettings()->setUpdateFields(true);
    // 设置默认的字体
    $word->getPhpWord()->setDefaultFontName("Times New Roman");
    $word->getPhpWord()->setDefaultFontSize($word->getFontSize("小四"));
    // 忽略错误
    $word->getPhpWord()->getSettings()->setHideGrammaticalErrors(true);
    $word->getPhpWord()->getSettings()->setHideSpellingErrors(true);
    // 正文样式  宋体小四  两端对齐  首行缩进2字符  行距固定20磅 段前、段后0.25行

    // 设置正文的字体和段落样式 宋体小四|两端对齐|首行缩进2字符|行距固定20磅|段前、段后0.25行
    $word->zw_font_style = $word->getFontStyle($fontName, $fontSize, false);
    $word->zw_para_style = $word->getParagraphStyle("both", $word->pointToTwip(20), ["firstLine" => $word->firstIndent($fontSize) * 2], false, false, null, $word->pointToLine(0.25, $fontSize), $word->pointToLine(0.25, $fontSize), false, "exact");
    $word->zw_font_bold_style = $word->getFontStyle($fontName, $fontSize, true);

    // 设置标题的字体和段落样式 宋体小四加粗|两端对齐|悬挂缩进2.34字符|多倍行距1.25倍|段前0.5行
    $word->bt_font_style = $word->getFontStyle($fontName, $fontSize, true);
    $word->bt_para_style = $word->getParagraphStyle("both", null, ["hanging" => $word->firstIndent($fontSize) * 2.34], false, false, 1.25, $word->pointToLine(0.5, $fontSize), $word->pointToLine(0.5, $fontSize));


    // 设置标题1的样式         宋体小四加粗|居中对齐|多倍行距1.25倍|段前、段后0.5行
    $word->getPhpWord()->addTitleStyle(1,
        $word->zw_font_bold_style,
        $word->getParagraphStyle("center", null, null, false, false, 1.25, $word->pointToLine(0.5, $fontSize), $word->pointToLine(0.5, $fontSize))
    );

    // 设置标题2的样式          宋体小四加粗|两端对齐|多倍行距0.2倍|段前、段后0.35行
    $word->getPhpWord()->addTitleStyle(2,
        $word->zw_font_bold_style,
        $word->getParagraphStyle("center", null, null, false, false, 1.25, $word->pointToLine(0.35, $fontSize), $word->pointToLine(0.35, $fontSize))
    );
    // 设置标题3的样式     宋体小四加粗|两端对齐|悬挂缩进2.34字符|多倍行距1.25倍|段前0.5行
    $word->getPhpWord()->addTitleStyle(3,
        $word->bt_font_style,
        $word->bt_para_style
    );
    // 添加列表样式
    $listTile = ["列表样式1", "列表样式2", "列表样式3", "列表样式4", "列表样式5", "列表样式6"];
    foreach ($listTile as $item) {
        $word->getPhpWord()->addNumberingStyle(
            $item,
            [
                'type' => 'multilevel',
                'levels' => [
                    ['format' => 'decimal', 'text' => '%1. ']
                ]
            ]
        );
    }
    // 设置表格的字体的样式
    $tableFont = "宋体";
    $tableFontSize = "五号";
    $word->table_font_style = $word->getFontStyle($tableFont, $tableFontSize);
    $word->table_bold_font_style = $word->getFontStyle($tableFont, $tableFontSize, true);
}

/**
 * 目录
 * @param MyPhpWord $word
 */
function category(MyPhpWord $word)
{
    $mlSection = $word->getPhpWord()->addSection();
    $mlSection->addText("目录", $word->zw_font_bold_style, $word->getParagraphStyle());
    $mlSection->addTOC(
        $word->zw_font_bold_style,
        [
            'tabLeader' => "dot"
        ]);
    $mlSection->addFooter()->addPreserveText('{PAGE}', "", ['align' => 'center']);
}

function testSelfDefine(MyPhpWord $word, \PhpOffice\PhpWord\Element\Section $section)
{
    $section->addListItem("列表1", 0, $word->bt_font_style, "列表样式1", $word->bt_para_style);
    $gsRun = $section->addTextRun($word->zw_para_style);
    $gsRun->addText("根据公司", $word->zw_font_style, $word->zw_para_style);
    $gsRun->addLink("附录", "①", $word->getFontStyle(null, null, false, "blue", true), null, true);
    $gsRun->addText("无换行拼接", $word->zw_font_style, $word->zw_para_style);
    $keyValue = [
        "名字" => 'name',
        "性别" => "sex",
        "年龄" => "age",
    ];
    // col为表头
    $word->tableColHeader(
        $keyValue,
        [
            "name" => "STAM.LIANG",
            "age" => "25",
            "sex" => "男",
        ],
        generateKeyValue($keyValue, 4, 10),
        generateKeyValue($keyValue, $word->getCellStyle("bdd6ee", "center"), $word->getCellStyle("ffffff", "center")),
        generateKeyValue($keyValue, $word->table_bold_font_style, $word->table_font_style),
        generateKeyValue($keyValue, $word->getParagraphStyle("center"), "left"), $section);

    $section->addTextBreak(1);
    // 列表2
    $section->addListItem("列表2", 0, $word->bt_font_style, "列表样式1", $word->bt_para_style);
//        $data = getStockholdersListByName($company, $enterprise_id, $tune_up_id = 0);
    $gdRun = $section->addTextRun($word->zw_para_style);
    $gdRun->addText("无换行拼接", $word->zw_font_style, $word->zw_para_style);
    $gdRun->addLink("附录", "①", $word->getFontStyle(null, null, false, "blue", true),
        null, true);
    $gdRun->addText("无换行拼接");

    $row_cn = ["编号", "姓名", "年龄", "性别"];
    $fields = ["sort_num", "name", "age", "sex"];
    $cellWidth = generateKeyValue(array_combine($row_cn, $fields), 3, 3);
    $cellWidth["序号"] = 1.5;
    $cellWidth["sort_num"] = 1.5;
    $row_data = [];
    for ($i = 0; $i < 5; $i++) {
        $row_data[$i]['sort_num'] = $i + 1;
        $row_data[$i]['name'] = "姓名" . $i;
        $row_data[$i]['age'] = "年龄" . $i;
        $row_data[$i]['sex'] = "性别" . $i;
    }
    // row为表头
    $word->tableRowHeader(1.5,
        $row_cn,
        $cellWidth,
        $row_data, $fields,
        generateKeyValue(array_combine($row_cn, $fields), $word->getCellStyle("bdd6ee", "center"), $word->getCellStyle("ffffff", "center")),
        generateKeyValue(array_combine($row_cn, $fields), $word->table_bold_font_style, $word->table_font_style),
        generateKeyValue(array_combine($row_cn, $fields), $word->getParagraphStyle("center"), $word->getParagraphStyle("center")),
        $section
    );
    $section->addTextBreak(1);

//    //  插图
    $section->addListItem("列表3", 0, $word->bt_font_style, "列表样式1", $word->bt_para_style);
    try {
        $word->checkImage("http://pic25.nipic.com/20121112/9252150_150552938000_2.jpg");
        $section->addImage("http://pic25.nipic.com/20121112/9252150_150552938000_2.jpg", $word->getImageStyle(null, null, 500, "square", "absolute", 0, 0));
    } catch (Exception $e) {
        print_r($e);
    }


    $section->addTextBreak(1);
}

function generateKeyValue(array $keyValue, $keySet, $valueSet)
{
    $gen = [];
    foreach ($keyValue as $key => $value) {
        $gen[$key] = $keySet;
        $gen[$value] = $valueSet;
    }
    return $gen;
}

setDocInfo($word);
initStyleSetting($word);
category($word);
$section = $word->getPhpWord()->addSection();
testSelfDefine($word, $section);


$flSection = $word->getPhpWord()->addSection();
$flSection->addTitle("附 录", 1);
$flTextRun = $flSection->addTextRun($word->zw_para_style);
$flTextRun->addBookmark("附录");
$flSection->addLink("http://cpquery.sipo.gov.cn/txnPantentInfoList.do?inner-flag:open-type=window&inner-flag:flowno=1494405123950", "1.参见中国及多国专利查询：".urlencode("http://cpquery.sipo.gov.cn/txnPantentInfoList.do?inner-flag:open-type=window&inner-flag:flowno=1494405123950"), $word->zw_font_style, $word->zw_para_style);
$flSection->addLink("http://www.ccopyright.com.cn/index.php?com=com_noticeQuery&method=softwareList&optionid=1221", "2.参见中国版权保护中心：".urlencode("http://www.ccopyright.com.cn/index.php?com=com_noticeQuery&method=softwareList&optionid=1221"), $word->zw_font_style, $word->zw_para_style);
$flSection->addLink("http://www.miitbeian.gov.cn/publish/query/indexFirst.action", "3.参见ICP/IP地址/域名信息备案管理系统：".urlencode("http://www.miitbeian.gov.cn/publish/query/indexFirst.action"), $word->zw_font_style, $word->zw_para_style);
$word->save();
