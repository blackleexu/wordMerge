<?php
/**
 * Created by PhpStorm.
 * User: black box
 * Date: 2019/9/17
 * Time: 12:32
 */

//word 合并  wordxml格式解析标签 参考 https://www.cnblogs.com/klbc/p/4005799.html
use Jupitern\DocxMerge\TbsZip;

include_once('tbszip.php');

mergeDoc('test2.docx','test1.docx');

function mergeDoc($path1,$path2){
    $zip = new TbsZip();

    // Open the first document
    $zip->Open($path2);
    $content1 = $zip->FileRead('word/document.xml');
    $zip->Close();

    // Extract the content of the first document
    $p = strpos($content1, '<w:body');
    if ($p===false){
        echo 'merge fail';
        exit("Tag </w:body> not found in document 2.");
    }
    $p = strpos($content1, '>', $p);
    $content1 = substr($content1, $p+1);
    $p = strpos($content1, '</w:body>');
    if ($p===false){
        echo 'merge fail';
        exit("Tag </w:body> not found in document 2.");
    } $content1 = substr($content1, 0, $p);

    // Insert into the second document
    $zip->Open($path1);
    $content2 = $zip->FileRead('word/document.xml');
    $p = strpos($content2, '</w:body>');

    //$content2 = '<w:p/>';换行符 如果需要的话可以加载$str查看效果

    //如果需要每个文件紧接在一起合并，不需要换页就去掉$str
    $str = '
    -<w:p>
    
    -<w:r>
    
    <w:br w:type="page"/>
    
    </w:r>
    
    </w:p>';

    if ($p===false){
        echo 'merge fail';
        exit("Tag </w:body> not found in document 1.");
    }
    $content2 = substr_replace($content2,$str. $content1, $p, 0);
    $zip->FileReplace('word/document.xml', $content2, TBSZIP_STRING);

    // Save the merge into a third file
    $result = 'merge'.time().'.docx';
    $zip->Flush(TBSZIP_FILE, $result);

    echo 'merge success path :'.__DIR__.'/'.$result;
}