// docx2html.js 20150305 陳爽 葉健欣
// 範例 node docx2html 或
//		node docx2html xlaw
// 說明 原始資料夾 \ksana2015\law6\raw\xlaw 或 \ksana2015\law6\raw\xlaw_1226
//		所有 .docx 擷取其中 text 及 link 在目的資料夾 law 產生對應 .html
//		若在 \ksana2015\law6\raw\xlaw_1226 遇同名 .docx 則直接讀取
// 注意 此程式需 docx2link.js 配合執行 以擷取 .docx 外部目標網址對照表
var fs=require("fs");
var inpdir=process.argv[2]||"xlaw" ;	// 預設 inpdir 為 xlaw, 因此 xlaw 可省略
inpdir='/ksana2015/law6/raw/'+inpdir;
console.log('原始夾',inpdir,fs.existsSync(inpdir));		// 原始資料夾 inpdir \ksana2015\law6\raw\xlaw
var lst=fs.readdirSync(inpdir).filter(function(fn){		// 蒐集原始資料夾中 .docx 檔名
	return fn.match(/^[^~$.]+\.docx$/);	
})
var pridir=inpdir+'_1226';
console.log('優先夾',pridir,fs.existsSync(pridir));		// 優先資料夾
var lst_1226=[];
if(fs.existsSync(pridir)){
	lst_1226=fs.readdirSync(pridir).filter(function(fn){// 蒐集優先資料夾中 .docx 檔名
		return fn.match(/^[^~$.]+\.docx$/);
	})
}
var m=inpdir.match(/\/x(.+?)$/);
if(!m){
	console.log('原始資料夾名稱',inpdir,'非以x起首');
	exit;
}
var outdir=m[1]
console.log('outdir',outdir);			// 目的資料夾 outdir law
var tstCount=0, fm=lst.length;
lst.forEach(function(filename,fi){
	var lawname=filename.match(/^(.+?)\.docx$/)[1];
	var pathname=inpdir+(lst_1226.indexOf(filename)>=0?'_1226':'')+'/'+filename; // 考慮優先資料夾
	var link=require("./docx2link")(pathname);		// 擷取 .docx 外部目標網址對照表
	var data=fs.readFileSync(pathname);				// 讀取 .docx
	var zip=require("jszip")(data);					// 產生對應 .zip
	var xml=zip.file("word/document.xml").asText(); // 取 .zip 內對應 .xml
	var msg='\033[0G'+fi+'/'+fm+' '+pathname+' ==> '+outdir+'\\'+lawname+'.html';
	process.stdout.write(msg);	
	var lines=[										// 蒐集要輸出到網頁的文字
		'<html><head><meta charset="utf-8"><title>'+lawname+'</title></head><body><pre>',
		msg
	];
	var line='', id='', tx='', instrText='', tagstack=[], fld=[''], pstyle=-1;
	var onTblOpen =function(node){ line+='<table>'; }
	var onTrOpen  =function(node){ line+='<tr>'; }
	var onTcOpen  =function(node){ line+='<td>'; }
	var onTcClose =function(node){ line+='</td>'; }
	var onTrClose =function(node){ line+='</tr>'; }
	var onTblClose=function(node){ lines.push(line+'</table>'), line=''; }
	var onP=function(node) {
		if (pstyle===2) replaceEntry();
		else if (pstyle===1) replaceChapter();
		line=line.replace(/\r?\n/,'');
		if(fld.length>2) anotherHyperLink();
		if(line) lines.push(line), line='';
		instrText=0, instrText="", pstyle=-1
	}
	var onPSyle=function(node) {
		pstyle=parseInt(node.attributes["w:val"]);
	}
	var anotherHyperLink=function(){
		var m=fld[1].match(/ HYPERLINK ( \\l )?"(.+?)" /);
		if(m){	var lnk=m[2];
				if(m[1]) lnk='#'+lnk; // local reference
				line=fld[0]+'<a href="'+lnk+'">'+fld[2]+'</a>'+(fld[3]||'');
				fld=[''];
		}
	}
	var replaceEntry=function(){
		var m=line.match(/第(.+?)條(（.+?）)/)
		if(m) { id=m[1],tx=m[2];
			line=line.replace(m[0],'<entry n="'+id+'">第'+id+'條'+tx+'</entry>');
		}
	}
	var replaceChapter=function(){
		var m=line.match(/第([^<>]+?)章([^<>]+)$/)
		if(m) { id=m[1],tx=m[2];
			line=line.replace(m[0],'<chapter n="'+id+'">第'+id+'章'+tx+'</chapter>');
		}
	}
	var replaceHyperLink=function(node){
		var lnk=node.attributes['w:anchor']
		if(lnk==undefined)
			lnk=link[node.attributes['r:id']];
		else
			lnk='#'+lnk;
		line=line.replace(hyperText,'<a href="'+lnk+'">'+hyperText+'</a>');
		hyperText="", hyperlink=0;
	}
	var insertHyperName=function(node){
		line+='<a name="'+node.attributes['w:name']+'"></a>';
	}
	var insertImage=function(node){
		line+='<img src="'+link[node.attributes['r:id']]+'"/>';
	}
	var hyperlink=0;
	var nodename; // 全域變數
	var strict= true;
	var parser= require("sax").parser(strict); // https://github.com/isaacs/sax-js
	parser.onclosetag = function (nodename) { var node=tagstack.pop();
			 if(nodename==="w:tbl"			) onTblClose		(node);
		else if(nodename==="w:tr"			) onTrClose 		(node);
		else if(nodename==="w:tc"			) onTcClose			(node);
		else if(nodename==="w:p"			) onP				(node);
		else if(nodename==="w:pStyle"		) onPSyle			(node);
		else if(nodename==='w:hyperlink'	) replaceHyperLink	(node);
		else if(nodename==='w:bookmarkStart') insertHyperName	(node);
		else if(nodename==='v:imagedata'	) insertImage		(node);
	};
	parser.onopentag= function (node) { nodename=node.name;
			 if(nodename==="w:tbl"			) onTblOpen			(node);
		else if(nodename==="w:tr" 			) onTrOpen			(node);
		else if(nodename==="w:tc" 			) onTcOpen			(node);
		else if(nodename==='w:hyperlink'	) hyperlink=1, hyperText='';
		else if(nodename==='w:fldChar'  	) fld.push('');
		tagstack.push(node);
	};
	parser.ontext=function(text){ text=text.replace(/>/g,'&gt;');
			 if (nodename==='w:t'			) fld[fld.length-1]+=text, line+=text;
		else if (nodename==='w:instrText'	) fld[fld.length-1]+=text;
		if (hyperlink) hyperText+=text;
	};
	parser.onend=function() {
		fs.writeFileSync(outdir+'/'+lawname+".html",lines.join('\n')+'</body></html>');
	}
	parser.write(xml).close();
});