// docx2link.js 20150305 陳爽 葉健欣
// 範例 link=require("docx2link")("高級中等教育法.docx");
// 說明 取出 高級中等教育法.docx 內 link Id 的外部目標網址資訊
// 運用 link[id] 即為對應的外部目標網址
var fs = require("fs");
function docxLink(filename){
	var data  = fs.readFileSync(filename); // 若依範例 即 讀取 高級中等教育法.docx
	var zip   = require("jszip")(data); // 相當於 產生 高級中等教育法.zip
	var rels  = zip.file("word/_rels/document.xml.rels").asText(); // zip 內 link Id 的 Target 資訊
	//var rels= fs.readFileSync("rels.xml","utf8"); // 直接取 實體 zip 內 事先更名之對應 xml 檔
	var link  = {}, strict= true; // strict=false for html-mode
	var parseLink= require("sax").parser(strict); // 預訂 parseLink 程序
	parseLink.onopentag=function(node){ // 檢視 "Relationship" 取 link Id 的 Target 資訊
		if (node.name==="Relationship") link[node.attributes.Id]=node.attributes.Target;
	};
	//parseLink.onerror=function(err){/* error happened */};
	//parseLink.ontext=function(t){/* t is the string of text */};
	//parseLink.onattribute=function(attr){/* attr has "name" and "value" */};
	//parseLink.onend=function(){/* parser stream is done, ready to have extra stuff written for it */};
	parseLink.write(rels).close(); // 執行 parseLink 程序 write(rels) 並 close()
	return link;
}
module.exports=docxLink; // 將 蒐集 link 的程序 docxLink 輸出 (link[Id] 即 Id 所對應的 Target)