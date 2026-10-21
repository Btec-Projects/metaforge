(function(){
  var $=function(id){return document.getElementById(id)};
  var zip=null,fileName="",orig={},coreXml="",appXml="",docXml="";
  var up=function(s){return s.charAt(0).toUpperCase()+s.slice(1)};

  // Theme toggle (remembered per viewer when storage is available)
  var root=document.documentElement;
  document.querySelectorAll("input[type=text]").forEach(function(i){i.dir="auto"});
  function view(v){if(["how","faq","credits","terms"].indexOf(v)<0)v="tool";$("tool").hidden=v!=="tool";["how","faq","credits","terms"].forEach(function(p){$("p_"+p).hidden=v!==p});["tool","how","faq","credits","terms"].forEach(function(p){$("tab_"+p).setAttribute("aria-current",v===p)});window.scrollTo(0,0)}
  ["tool","how","faq","credits","terms"].forEach(function(p){$("tab_"+p).addEventListener("click",function(){view(p);location.hash=p})});
  document.querySelectorAll("[data-go]").forEach(function(b){b.addEventListener("click",function(){view(b.dataset.go);location.hash=b.dataset.go})});
  
  window.addEventListener("hashchange",function(){view(location.hash.slice(1))});
  view(location.hash.slice(1));
  $("presetSafe").addEventListener("click",function(){["author","lastBy","company","subject","keywords","desc"].forEach(function(k){$(k).value=""});$("accept").checked=true;renderDiff();msg("Share-safely preset applied. Review the changes, then save.","ok")});
  async function inspect(){
    var box=$("inspect");box.textContent="";var cx="",cf=zip.file("word/comments.xml");if(cf)cx=await cf.async("string");
    var n=Object.keys(zip.files);
    var items=[["Tracked changes",(docXml.match(/<w:(?:ins|del)\b[^>]*[^\/]>/g)||[]).length,1],["Comments",(cx.match(/<w:comment\b/g)||[]).length,1],["Images",n.filter(function(k){return /^word\/media\//.test(k)}).length,0],["Hyperlinks",(docXml.match(/<w:hyperlink\b/g)||[]).length,0],["Headers and footers",n.filter(function(k){return /^word\/(header|footer)\d*\.xml$/.test(k)}).length,0],["Custom properties",zip.file("docProps/custom.xml")?"yes":"none",0]];
    items.forEach(function(it){var s=document.createElement("span");s.textContent=it[0]+": "+it[1];if(it[2]&&it[1])s.className="hit";box.appendChild(s)});
  }
  async function sample(){
    var X='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',R="http://schemas.openxmlformats.org/",z=new JSZip();
    z.file("[Content_Types].xml",X+'<Types xmlns="'+R+'package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>');
    z.file("_rels/.rels",X+'<Relationships xmlns="'+R+'package/2006/relationships"><Relationship Id="rId1" Type="'+R+'officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="'+R+'package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="'+R+'officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>');
    z.file("word/document.xml",X+'<w:document xmlns:w="'+R+'wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Moreover, we utilize a approach to delve into the the data , in order to learn.</w:t></w:r></w:p><w:p><w:r><w:t>مَرْحَبًا، هذا نصٌّ تجريبيّ.</w:t></w:r></w:p></w:body></w:document>');
    z.file("docProps/core.xml",X+'<cp:coreProperties xmlns:cp="'+R+'package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Sample report</dc:title><dc:creator>Sample Author</dc:creator><cp:lastModifiedBy>Sample Author</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">2026-01-05T09:00:00Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2026-01-05T10:00:00Z</dcterms:modified></cp:coreProperties>');
    z.file("docProps/app.xml",X+'<Properties xmlns="'+R+'officeDocument/2006/extended-properties"><TotalTime>45</TotalTime><Company>Sample Co</Company></Properties>');
    load(new File([await z.generateAsync({type:"blob"})],"sample.docx"));
  }
  $("trySample").addEventListener("click",sample);
  function saveOpts(){try{localStorage.setItem("mf-opts",JSON.stringify([].slice.call(document.querySelectorAll("#rules input:checked,#chars input:checked")).map(function(i){return i.value})))}catch(_){}}
  function restoreOpts(){try{var s=JSON.parse(localStorage.getItem("mf-opts")||"[]");document.querySelectorAll("#rules input,#chars input").forEach(function(i){i.checked=s.indexOf(i.value)>-1});renderDiff()}catch(_){}}
  try{var t=localStorage.getItem("mf-theme");if(t)root.dataset.theme=t}catch(_){}
  $("theme").addEventListener("click",function(){
    var dark=root.dataset.theme?root.dataset.theme==="dark":matchMedia("(prefers-color-scheme:dark)").matches;
    root.dataset.theme=dark?"light":"dark";
    try{localStorage.setItem("mf-theme",root.dataset.theme)}catch(_){}
  });

  var SO="(^|[.!?]\\s+)";
  var RULES=[
    {id:"openers",label:"Opening connectors",re:new RegExp(SO+"(?:Moreover|Furthermore|Additionally|In conclusion|In summary|Overall),\\s+(\\S)","g"),rep:function(m,a,b){return a+up(b)}},
    {id:"noting",label:"“It is important to note that”",
      re:/(^|[.!?]\s+)?[Ii]t(?: is|’s|'s) (?:important|worth|crucial) (?:to note|noting|to mention) that\s+(\S)/g,
      rep:function(m,a,b){return a?a+up(b):b}},
    {id:"opening",label:"Cliché openers",re:new RegExp(SO+"In (?:today’s|today's) (?:fast-paced |digital |modern )?world,\\s+(\\S)","g"),rep:function(m,a,b){return a+up(b)}},
    {id:"delve",label:"“delve into” → “look into”",
      re:/\b(D|d)elve(s|d)?\b( into)?/g,rep:function(m,c,s,i){var w=(c==="D"?"L":"l")+"ook"+(s==="s"?"s":s==="d"?"ed":"");return w+(i?" into":"")}},
    {id:"utilize",label:"“utilize” → “use”",
      re:/\b(U|u)tili[sz](e|es|ed|ing)\b/g,
      rep:function(m,c,f){return (c==="U"?"U":"u")+{e:"se",es:"ses",ed:"sed",ing:"sing"}[f]}},
    {id:"tapestry",label:"“rich tapestry of” → “mix of”",re:/\b(?:rich )?tapestry of\b/g,rep:function(){return "mix of"}},
    {id:"testament",label:"“a testament to” → “proof of”",re:/\b(a|A) testament to\b/g,rep:function(m,c){return c==="A"?"Proof of":"proof of"}},
    {id:"landscape",label:"“navigate the complexities of” → “handle”",re:/\b(N|n)avigat(?:e|ing) the complexities of\b/g,rep:function(m,c){return (c==="N"?"H":"h")+"andle"}},
    {id:"inorder",label:"“in order to” → “to”",re:/\b(I|i)n order to\b/g,rep:function(m,c){return c==="I"?"To":"to"}},
    {id:"dueto",label:"“due to the fact that” → “because”",re:/\b(D|d)ue to the fact that\b/g,rep:function(m,c){return c==="D"?"Because":"because"}},
    {id:"fillers",label:"Filler asides (“Needless to say,” “At the end of the day,”)",re:new RegExp(SO+"(?:Needless to say|At the end of the day|It goes without saying),\\s+(\\S)","g"),rep:function(m,a,b){return a+up(b)}},
    {id:"leverage",label:"“leverage” → “use”",re:/\b(L|l)everag(e|es|ed|ing)\b/g,rep:function(m,c,f){return (c==="L"?"U":"u")+{e:"se",es:"ses",ed:"sed",ing:"sing"}[f]}},
    {id:"crucial",label:"“plays a crucial role in” → “is key to”",re:/\bplays an? (?:crucial|vital|pivotal) role in\b/g,rep:function(){return "is key to"}},
    {id:"g_space",label:"Grammar: double spaces",re:/(\S) {2,}(?=\S)/g,rep:function(m,a){return a+" "}},
    {id:"g_punct",label:"Grammar: space before , . ; : ! ?",re:/(\S) +([,.;:!?])/g,rep:function(m,a,p){return a+p}},
    {id:"g_repeat",label:"Grammar: repeated words (“the the”)",re:/\b([A-Za-z]{2,}) \1\b/gi,rep:function(m,w){return /^(had|that)$/i.test(w)?m:w}},
    {id:"g_an",label:"Grammar: “a” before a vowel → “an”",re:/\b(a) (?=(?!uni|use|usu|eu|one|once|ubi|uti)[aeiou]\w)/gi,rep:function(m,c){return c+"n "}}
  ];
  function acceptAll(x){return x.replace(/<w:del\b[^>]*[^\/]>[\s\S]*?<\/w:del>/g,"").replace(/<\/?w:ins\b[^>]*>/g,"")}
  var KEEP=function(m,d){return d?m:""};
  var CHARS=[
    {id:"c_period",label:"Remove periods ( . ) — decimals like 3.14 stay",re:/(\d)\.(?=\d)|\./g,rep:KEEP},
    {id:"c_comma",label:"Remove commas ( , and ، ) — numbers like 1,000 stay",re:/(\d),(?=\d)|[,،]/g,rep:KEEP},
    {id:"c_marks",label:"Remove Arabic diacritics ( َ ً ُ ٌ ِ ٍ ّ ْ )",re:/[\u064B-\u0652]/g,rep:function(){return ""}},
    {id:"c_tatweel",label:"Remove Arabic tatweel stretching ( ـ )",re:/\u0640/g,rep:function(){return ""}},
    {id:"c_quotes",label:"Straighten curly quotes ( “ ” ‘ ’ )",re:/[\u201C\u201D]|[\u2018\u2019]/g,rep:function(m){return /[\u201C\u201D]/.test(m)?'"':"'"}}
  ];
  function applyRules(text,ids,counter){
    RULES.concat(CHARS).forEach(function(r){
      if(ids.indexOf(r.id)<0)return;
      text=text.replace(r.re,function(){counter[r.id]=(counter[r.id]||0)+1;return r.rep.apply(null,arguments)});
    });
    return text;
  }
  function mapRuns(xml,fn){
    return xml.replace(/(<w:t(?:\s[^>]*)?>)([^<]*)(<\/w:t>)/g,function(_,o,t,c){return o+esc(fn(unesc(t)))+c});
  }
  function msg(t,c){var m=$("msg");m.textContent=t;m.className=c||""}
  function esc(s){return s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")}
  function unesc(s){return s.replace(/&lt;/g,"<").replace(/&gt;/g,">").replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&amp;/g,"&")}
  function getTag(xml,tag){var m=xml.match(new RegExp("<"+tag+"(?:\\s[^>]*)?>([\\s\\S]*?)</"+tag+">"));return m?unesc(m[1]):""}
  function setTag(xml,tag,val,attrs){
    var re=new RegExp("<"+tag+"((?:\\s[^>]*)?)>[\\s\\S]*?</"+tag+">"),safe=esc(val);
    if(re.test(xml))return xml.replace(re,function(_,a){return "<"+tag+a+">"+safe+"</"+tag+">"});
    var self=new RegExp("<"+tag+"(?:\\s[^>]*)?/>"),el="<"+tag+(attrs||"")+">"+safe+"</"+tag+">";
    if(self.test(xml))return xml.replace(self,el);
    return xml.replace(/<\/(cp:coreProperties|Properties)>\s*$/,el+"</$1>");
  }
  var pad=function(n){return String(n).padStart(2,"0")};
  function toLocalInput(d){return d.getFullYear()+"-"+pad(d.getMonth()+1)+"-"+pad(d.getDate())+"T"+pad(d.getHours())+":"+pad(d.getMinutes())+":"+pad(d.getSeconds())}
  function fromIso(iso){var d=new Date(iso);return isNaN(d)?"":toLocalInput(d)}
  function toIso(local){var d=new Date(local);return isNaN(d)?"":d.toISOString().replace(/\.\d{3}Z$/,"Z")}
  function setTime(total){total=Math.max(0,total|0);$("hours").value=Math.floor(total/60);$("mins").value=total%60}
  function getTime(){return (parseInt($("hours").value,10)||0)*60+(parseInt($("mins").value,10)||0)}
  var TEXT={author:"author",lastBy:"lastBy",title:"title",company:"company",subject:"subject",keywords:"keywords",desc:"desc"};
  function fill(){
    Object.keys(TEXT).forEach(function(k){$(k).value=orig[k]});
    $("created").value=fromIso(orig.created);$("modified").value=fromIso(orig.modified);
    setTime(parseInt(orig.minutes,10)||0);
    $("sync").checked=false;
    document.querySelectorAll("#rules input,#chars input,#accept").forEach(function(i){i.checked=false});renderDiff();
  }
  function buildRules(){
    var counter={};mapRuns(docXml,function(t){applyRules(t,RULES.concat(CHARS).map(function(r){return r.id}),counter);return t});
    $("rules").innerHTML="";$("chars").innerHTML="";
    RULES.concat(CHARS).forEach(function(r){
      var l=document.createElement("label");l.className="check";
      var i=document.createElement("input");i.type="checkbox";i.value=r.id;
      var s=document.createElement("span");
      s.innerHTML="<b></b><span class='count'></span>";
      s.firstChild.textContent=r.label;s.lastChild.textContent=(counter[r.id]||0)+" found";
      l.appendChild(i);l.appendChild(s);(r.id.slice(0,2)==="c_"?$("chars"):$("rules")).appendChild(l);
    });
  }
  function fmtSize(b){return b>1048576?(b/1048576).toFixed(1)+" MB":Math.max(1,Math.round(b/1024))+" KB"}

  async function load(f){
    if(!f)return;msg("Reading file…");
    if(!/\.docx$/i.test(f.name)){msg("That file is not a .docx. Choose a Word document saved as .docx.","err");return}
    try{
      zip=await JSZip.loadAsync(await f.arrayBuffer());
      var core=zip.file("docProps/core.xml"),app=zip.file("docProps/app.xml"),doc=zip.file("word/document.xml");
      if(!core){msg("This file has no document properties, so there is nothing to edit.","err");zip=null;return}
      coreXml=await core.async("string");appXml=app?await app.async("string"):"";docXml=doc?await doc.async("string"):"";
      fileName=f.name;
      orig={author:getTag(coreXml,"dc:creator"),lastBy:getTag(coreXml,"cp:lastModifiedBy"),
        created:getTag(coreXml,"dcterms:created"),modified:getTag(coreXml,"dcterms:modified"),
        minutes:getTag(appXml,"TotalTime")||"0",title:getTag(coreXml,"dc:title"),company:getTag(appXml,"Company"),
        subject:getTag(coreXml,"dc:subject"),keywords:getTag(coreXml,"cp:keywords"),desc:getTag(coreXml,"dc:description")};
      $("fname").textContent=fileName;
      var w=getTag(appXml,"Words"),p=getTag(appXml,"Pages");
      $("fmeta").textContent=[fmtSize(f.size),p&&p+" pages",w&&w+" words"].filter(Boolean).join(" · ");
      buildRules();fill();restoreOpts();inspect();
      $("hours").disabled=$("mins").disabled=!app;
      $("editor").hidden=false;
      msg("File loaded. Change what you need, then save.","ok");
    }catch(e){zip=null;$("editor").hidden=true;msg("Couldn’t read that file. It may be damaged or password-protected.","err")}
  }

  function fmtT(t){return Math.floor(t/60)+"h "+(t%60)+"m"}
  function fmtD(v){var d=new Date(v);return isNaN(d)?"(not set)":d.toLocaleString()}
  function renderDiff(){
    if(!zip)return;
    var rows=[];
    function add(l,a,b){if(a!==b)rows.push([l,a||"(empty)",b||"(empty)"])}
    add("Author",orig.author,$("author").value);
    add("Last modified by",orig.lastBy,$("lastBy").value);
    add("Title",orig.title,$("title").value);
    add("Company",orig.company,$("company").value);
    add("Subject",orig.subject,$("subject").value);
    add("Keywords",orig.keywords,$("keywords").value);
    add("Comments",orig.desc,$("desc").value);
    add("Created",fmtD(orig.created),fmtD($("created").value));
    add("Last modified",fmtD(orig.modified),fmtD($("modified").value));
    add("Editing time",fmtT(parseInt(orig.minutes,10)||0),fmtT(getTime()));
    var n=document.querySelectorAll("#rules input:checked,#chars input:checked,#accept:checked").length;
    if(n)rows.push(["Text cleanups",n+" rule"+(n>1?"s":"")+" selected","applied on save"]);
    var box=$("diff");box.textContent="";
    if(!rows.length){box.textContent="Nothing changed yet."}
    else{var t=document.createElement("table");
      rows.forEach(function(r){var tr=t.insertRow();r.forEach(function(v){tr.insertCell().textContent=v})});
      box.appendChild(t)}
    var c=new Date($("created").value),m=new Date($("modified").value);
    $("warn").textContent=(!isNaN(c)&&!isNaN(m)&&m<c)?"Heads up: the last modified date is earlier than the created date.":"";
  }
  $("editor").addEventListener("input",function(e){renderDiff();if(e.target.closest("#rules,#chars"))saveOpts()});
  $("editor").addEventListener("click",function(){setTimeout(function(){renderDiff();saveOpts()},0)});
  $("clearPeople").addEventListener("click",function(){["author","lastBy","company","subject","keywords","desc"].forEach(function(k){$(k).value=""});renderDiff()});
  function setAll(v){document.querySelectorAll("#rules input").forEach(function(i){i.checked=v});renderDiff()}
  $("allRules").addEventListener("click",function(){setAll(true)});
  $("noRules").addEventListener("click",function(){setAll(false)});
  $("file").addEventListener("change",function(e){load(e.target.files[0])});
  var drop=$("drop");
  ["dragenter","dragover"].forEach(function(t){drop.addEventListener(t,function(e){e.preventDefault();drop.classList.add("over")})});
  ["dragleave","drop"].forEach(function(t){drop.addEventListener(t,function(e){e.preventDefault();drop.classList.remove("over")})});
  drop.addEventListener("drop",function(e){load(e.dataTransfer.files[0])});
  $("preview").addEventListener("click",function(){
    var ids=[].slice.call(document.querySelectorAll("#rules input:checked,#chars input:checked")).map(function(i){return i.value}),out=[],box=$("prev");
    box.textContent="";
    if(!zip||!ids.length){box.textContent="Select at least one cleanup to preview.";return}
    mapRuns(docXml,function(t){var n=applyRules(t,ids,{});if(n!==t&&out.length<5)out.push([t,n]);return t});
    if(!out.length){box.textContent="No text would change.";return}
    out.forEach(function(p){var d=document.createElement("p");d.className="fmeta";d.dir="auto";d.textContent=p[0].slice(0,120)+"  →  "+p[1].slice(0,120);box.appendChild(d)});
  });
  $("reset").addEventListener("click",function(){fill();msg("Reset to the original values.","")});
  $("sync").addEventListener("change",function(e){if(e.target.checked)$("lastBy").value=$("author").value});
  $("author").addEventListener("input",function(){if($("sync").checked)$("lastBy").value=$("author").value});
  document.querySelectorAll("[data-now]").forEach(function(b){b.addEventListener("click",function(){$(b.dataset.now).value=toLocalInput(new Date())})});
  $("afterCreated").addEventListener("click",function(){$("modified").value=$("created").value});
  document.querySelectorAll("[data-add]").forEach(function(b){b.addEventListener("click",function(){setTime(getTime()+parseInt(b.dataset.add,10))})});
  $("zero").addEventListener("click",function(){setTime(0)});
  [$("hours"),$("mins")].forEach(function(i){i.addEventListener("change",function(){setTime(getTime())})});
  document.addEventListener("keydown",function(e){
    if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==="s"){e.preventDefault();if(zip)$("editor").requestSubmit()}
  });

  $("editor").addEventListener("submit",async function(e){
    e.preventDefault();if(!zip)return;
    var cIso=toIso($("created").value),mIso=toIso($("modified").value);
    if(!cIso||!mIso){msg("Enter both a created date and a last modified date.","err");return}
    try{
      var core=setTag(coreXml,"dc:creator",$("author").value);
      core=setTag(core,"cp:lastModifiedBy",$("lastBy").value);core=setTag(core,"dc:title",$("title").value);
      core=setTag(core,"dc:subject",$("subject").value);core=setTag(core,"cp:keywords",$("keywords").value);core=setTag(core,"dc:description",$("desc").value);
      var dt=' xsi:type="dcterms:W3CDTF"';
      core=setTag(core,"dcterms:created",cIso,dt);core=setTag(core,"dcterms:modified",mIso,dt);
      zip.file("docProps/core.xml",core);
      if(appXml){var ax=setTag(appXml,"Company",$("company").value);if(!$("hours").disabled)ax=setTag(ax,"TotalTime",String(getTime()));zip.file("docProps/app.xml",ax)}
      var ids=[].slice.call(document.querySelectorAll("#rules input:checked,#chars input:checked")).map(function(i){return i.value});
      var changed=0;
      var acc=$("accept").checked;if((ids.length||acc)&&docXml){
        var c={};zip.file("word/document.xml",mapRuns(acc?acceptAll(docXml):docXml,function(t){return applyRules(t,ids,c)}));var extra=zip.file(/^word\/(header|footer|footnotes|endnotes)\d*\.xml$/);for(var q=0;q<extra.length;q++){var ex=await extra[q].async("string");zip.file(extra[q].name,mapRuns(ex,function(t){return applyRules(t,ids,c)}))}
        Object.keys(c).forEach(function(k){changed+=c[k]});
      }
      var blob=await zip.generateAsync({type:"blob",compression:"DEFLATE",mimeType:"application/vnd.openxmlformats-officedocument.wordprocessingml.document"});
      var out=fileName.replace(/\.docx$/i,"")+"-edited.docx";
      var url=URL.createObjectURL(blob),a=document.createElement("a");
      a.href=url;a.download=out;document.body.appendChild(a);a.click();a.remove();
      window.setTimeout(function(){URL.revokeObjectURL(url)},1000);
      msg("Saved "+out+(ids.length?" ("+changed+" phrase"+(changed===1?"":"s")+" changed).":"."),"ok");
    }catch(err){
      if(err&&err.code==="declined")msg("Save cancelled.","");
      else msg("Couldn’t save the file: "+((err&&err.message)||"unknown error"),"err");
    }
  });
})();
