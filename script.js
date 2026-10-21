(function(){
  var $=function(id){return document.getElementById(id)};
  var zip=null,fileName="",orig={},coreXml="",appXml="",docXml="",sourceFile=null,batchFiles=[];
  var ruleCounts={},privacyFindings={},undoHistory=[],lastState=null;
  var up=function(s){return s.charAt(0).toUpperCase()+s.slice(1)};

  var root=document.documentElement;
  document.querySelectorAll("input[type=text]").forEach(function(i){i.dir="auto"});
  function view(v){if(["how","faq","credits","terms"].indexOf(v)<0)v="tool";$("tool").hidden=v!=="tool";["how","faq","credits","terms"].forEach(function(p){$("p_"+p).hidden=v!==p});["tool","how","faq","credits","terms"].forEach(function(p){$("tab_"+p).setAttribute("aria-current",v===p)});window.scrollTo(0,0)}
  ["tool","how","faq","credits","terms"].forEach(function(p){$("tab_"+p).addEventListener("click",function(){view(p);location.hash=p})});
  document.querySelectorAll("[data-go]").forEach(function(b){b.addEventListener("click",function(){view(b.dataset.go);location.hash=b.dataset.go})});
  
  window.addEventListener("hashchange",function(){view(location.hash.slice(1))});
  view(location.hash.slice(1));
  $("presetSafe").addEventListener("click",function(){rememberUndo();["author","lastBy","company","subject","keywords","desc"].forEach(function(k){$(k).value=""});$("accept").checked=true;$("removeComments").checked=true;$("removeCustom").checked=true;renderDiff();msg("Privacy preset applied. Review the permanent removals before saving.","ok")});
  async function inspect(){
    var box=$("inspect");box.textContent="";
    var names=Object.keys(zip.files),xmlFiles=names.filter(function(n){return /\.(?:xml|rels)$/i.test(n)&&!zip.files[n].dir});
    var xmlParts=await Promise.all(xmlFiles.map(function(n){return zip.file(n).async("string").then(function(text){return {name:n,text:text}})}));
    var allXml=xmlParts.map(function(p){return p.text}).join("\n");
    var rels=xmlParts.filter(function(p){return /\.rels$/i.test(p.name)}).map(function(p){return p.text}).join("\n");
    var commentFile=zip.file("word/comments.xml"),comments=commentFile?await commentFile.async("string"):"";
    privacyFindings={
      personal:Object.keys(orig).filter(function(k){return ["author","lastBy","company","subject","keywords","desc"].indexOf(k)>-1&&orig[k]}).length,
      comments:(comments.match(/<w:comment\b/g)||[]).length,
      tracked:(allXml.match(/<w:(?:ins|del)\b/g)||[]).length,
      custom:zip.file("docProps/custom.xml")?(zip.file("docProps/custom.xml").async("string").then(function(text){return (text.match(/<property\b/g)||[]).length})):Promise.resolve(0),
      external:(rels.match(/\bTargetMode=["']External["']/g)||[]).length,
      embedded:names.filter(function(n){return /^word\/(embeddings|activeX)\//i.test(n)}).length,
      media:names.filter(function(n){return /^word\/media\//i.test(n)}).length
    };
    privacyFindings.custom=await privacyFindings.custom;
    var items=[["Personal metadata",privacyFindings.personal,1],["Comments",privacyFindings.comments,1],["Tracked changes",privacyFindings.tracked,1],["External links",privacyFindings.external,0],["Embedded files",privacyFindings.embedded,0],["Images",privacyFindings.media,0],["Custom properties",privacyFindings.custom,1]];
    items.forEach(function(it){var s=document.createElement("span");s.textContent=it[0]+": "+it[1];if(it[2]&&it[1])s.className="hit";box.appendChild(s)});
    renderReview();
  }
  async function sample(){
    var X='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',R="http://schemas.openxmlformats.org/",z=new JSZip();
    z.file("[Content_Types].xml",X+'<Types xmlns="'+R+'package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/><Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/></Types>');
    z.file("_rels/.rels",X+'<Relationships xmlns="'+R+'package/2006/relationships"><Relationship Id="rId1" Type="'+R+'officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="'+R+'package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="'+R+'officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId4" Type="'+R+'officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/></Relationships>');
    z.file("word/_rels/document.xml.rels",X+'<Relationships xmlns="'+R+'package/2006/relationships"><Relationship Id="rId1" Type="'+R+'wordprocessingml/2006/relationships/comments" Target="comments.xml"/></Relationships>');
    z.file("word/document.xml",X+'<w:document xmlns:w="'+R+'wordprocessingml/2006/main"><w:body><w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Moreover, we utilize a approach to delve into the the data , in order to learn.</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p><w:p><w:del w:id="1" w:author="Sample Author"><w:r><w:delText>Old sample wording</w:delText></w:r></w:del><w:ins w:id="2" w:author="Sample Author"><w:r><w:t>Reviewed sample wording</w:t></w:r></w:ins></w:p><w:p><w:r><w:t>مَرْحَبًا، هذا نصٌّ تجريبيّ.</w:t></w:r></w:p></w:body></w:document>');
    z.file("word/comments.xml",X+'<w:comments xmlns:w="'+R+'wordprocessingml/2006/main"><w:comment w:id="0" w:author="Sample Reviewer"><w:p><w:r><w:t>Private review note (sample)</w:t></w:r></w:p></w:comment></w:comments>');
    z.file("docProps/core.xml",X+'<cp:coreProperties xmlns:cp="'+R+'package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Sample report</dc:title><dc:creator>Sample Author</dc:creator><cp:lastModifiedBy>Sample Author</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">2026-01-05T09:00:00Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2026-01-05T10:00:00Z</dcterms:modified></cp:coreProperties>');
    z.file("docProps/app.xml",X+'<Properties xmlns="'+R+'officeDocument/2006/extended-properties"><TotalTime>45</TotalTime><Company>Sample Co</Company></Properties>');
    z.file("docProps/custom.xml",X+'<Properties xmlns="'+R+'officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="SampleCode"><vt:lpwstr>DEMO-ONLY</vt:lpwstr></property></Properties>');
    var sampleFile=new File([await z.generateAsync({type:"blob"})],"sample.docx");
    batchFiles=[sampleFile];refreshBatch();load(sampleFile);
  }
  $("trySample").addEventListener("click",sample);
  function saveOpts(){try{localStorage.setItem("mf-opts",JSON.stringify([].slice.call(document.querySelectorAll("#rules input:checked,#chars input:checked")).map(function(i){return i.value})))}catch(_){}}
  function restoreOpts(){try{var s=JSON.parse(localStorage.getItem("mf-opts")||"[]");document.querySelectorAll("#rules input,#chars input").forEach(function(i){i.checked=s.indexOf(i.value)>-1});renderDiff();syncUndoState()}catch(_){}}
  try{
    var theme=localStorage.getItem("mf-theme")||"system";
    if(["system","light","dark"].indexOf(theme)<0)theme="system";
    root.dataset.theme=theme;$("themeMode").value=theme;
  }catch(_){root.dataset.theme="system"}
  $("themeMode").addEventListener("change",function(){
    root.dataset.theme=$("themeMode").value;
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
  function acceptAll(x){return x.replace(/<w:del\b[^>]*\/>/g,"").replace(/<w:del\b[^>]*>[\s\S]*?<\/w:del>/g,"").replace(/<\/?w:ins\b[^>]*>/g,"")}
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
  function formSnapshot(){
    var values={};
    document.querySelectorAll("#editor input:not([type=file])").forEach(function(i){
      values[i.id]=i.type==="checkbox"?i.checked:i.value;
    });
    return values;
  }
  function sameSnapshot(a,b){return JSON.stringify(a)===JSON.stringify(b)}
  function updateUndoButton(){$("undo").disabled=!undoHistory.length}
  function resetUndo(){undoHistory=[];lastState=formSnapshot();updateUndoButton()}
  function rememberUndo(){
    var current=formSnapshot();
    if(lastState&&!sameSnapshot(current,lastState))undoHistory.push(lastState);
    else undoHistory.push(current);
    if(undoHistory.length>40)undoHistory.shift();
    lastState=current;updateUndoButton();
  }
  function syncUndoState(){lastState=formSnapshot();updateUndoButton()}
  function restoreSnapshot(values){
    Object.keys(values).forEach(function(id){
      var el=$(id);if(!el)return;
      if(el.type==="checkbox")el.checked=values[id];else el.value=values[id];
    });
    renderDiff();saveOpts();syncUndoState();
  }
  function fill(preserveUndo){
    Object.keys(TEXT).forEach(function(k){$(k).value=orig[k]});
    $("created").value=fromIso(orig.created);$("modified").value=fromIso(orig.modified);
    setTime(parseInt(orig.minutes,10)||0);
    $("sync").checked=false;
    document.querySelectorAll("#rules input,#chars input,#accept,#removeComments,#removeCustom").forEach(function(i){i.checked=false});
    renderDiff();if(!preserveUndo)resetUndo();else syncUndoState();
  }
  function buildRules(){
    var counter={};mapRuns(docXml,function(t){applyRules(t,RULES.concat(CHARS).map(function(r){return r.id}),counter);return t});
    ruleCounts=counter;
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

  function refreshBatch(){
    var panel=$("batchPanel"),list=$("batchList");
    list.textContent="";
    batchFiles.forEach(function(f){var li=document.createElement("li");li.textContent=f.name+" · "+fmtSize(f.size);list.appendChild(li)});
    panel.hidden=batchFiles.length<2;
    $("batchSave").hidden=batchFiles.length<2;
    $("batchSave").textContent=batchFiles.length>1?"Save "+batchFiles.length+" files as ZIP":"Save all as ZIP";
  }
  async function loadFiles(files){
    var selected=[].slice.call(files||[]),valid=selected.filter(function(f){return /\.docx$/i.test(f.name)});
    if(!valid.length){msg("Choose one or more Word documents saved as .docx.","err");return}
    batchFiles=valid;refreshBatch();
    var rejected=selected.length-valid.length;
    await load(valid[0]);
    if(rejected)msg((rejected===1?"Skipped 1 non-.docx file. ":"Skipped "+rejected+" non-.docx files. ")+"The first .docx is open; batch save uses these settings for all selected files.","ok");
    else if(valid.length>1)msg(valid.length+" files ready. The displayed settings will be applied to every file in the batch.","ok");
  }
  async function load(f){
    if(!f)return;$("saveAgain").hidden=true;msg("Reading file…");
    try{
      zip=await JSZip.loadAsync(await f.arrayBuffer());
      var core=zip.file("docProps/core.xml"),app=zip.file("docProps/app.xml"),doc=zip.file("word/document.xml");
      if(!core){msg("This file has no document properties, so there is nothing to edit.","err");zip=null;$("editor").hidden=true;return}
      sourceFile=f;
      coreXml=await core.async("string");appXml=app?await app.async("string"):"";docXml=doc?await doc.async("string"):"";
      fileName=f.name;
      orig={author:getTag(coreXml,"dc:creator"),lastBy:getTag(coreXml,"cp:lastModifiedBy"),
        created:getTag(coreXml,"dcterms:created"),modified:getTag(coreXml,"dcterms:modified"),
        minutes:getTag(appXml,"TotalTime")||"0",title:getTag(coreXml,"dc:title"),company:getTag(appXml,"Company"),
        subject:getTag(coreXml,"dc:subject"),keywords:getTag(coreXml,"cp:keywords"),desc:getTag(coreXml,"dc:description")};
      $("fname").textContent=fileName;
      var w=getTag(appXml,"Words"),p=getTag(appXml,"Pages");
      $("fmeta").textContent=[fmtSize(f.size),p&&p+" pages",w&&w+" words"].filter(Boolean).join(" · ");
      buildRules();fill();restoreOpts();await inspect();
      $("hours").disabled=$("mins").disabled=!app;
      $("editor").hidden=false;
      refreshBatch();
      msg(batchFiles.length>1?batchFiles.length+" files loaded. Current settings apply to the whole batch.":"File loaded. Change what you need, then save.","ok");
    }catch(e){zip=null;$("editor").hidden=true;msg("Couldn’t read that file. It may be damaged or password-protected.","err")}
  }

  function fmtT(t){return Math.floor(t/60)+"h "+(t%60)+"m"}
  function fmtD(v){var d=new Date(v);return isNaN(d)?"(not set)":d.toLocaleString()}
  function selectedRules(){return [].slice.call(document.querySelectorAll("#rules input:checked,#chars input:checked")).map(function(i){return i.value})}
  function renderReview(){
    if(!zip)return;
    var box=$("privacyReview");box.textContent="";
    var personalNames=["Author","Last modified by","Company","Subject","Keywords","Comments property"];
    var personalIds=["author","lastBy","company","subject","keywords","desc"];
    var presentPersonal=personalIds.reduce(function(found,id,index){if($(id).value)found.push(personalNames[index]);return found},[]);
    var personalCount=presentPersonal.length;
    var items=[
      ["Personal metadata",personalCount,false],
      ["Document comments",privacyFindings.comments,$("removeComments").checked],
      ["Tracked changes",privacyFindings.tracked,$("accept").checked],
      ["Custom properties",privacyFindings.custom,$("removeCustom").checked],
      ["External links",privacyFindings.external,false],
      ["Embedded files",privacyFindings.embedded,false],
      ["Images",privacyFindings.media,false]
    ];
    items.forEach(function(item){
      var li=document.createElement("li"),count=item[1]||0;
      li.className=count?(item[2]?"review-clean":"review-found"):"review-clear";
      var detail=count?(item[2]?(item[0]==="Tracked changes"?"will be accepted; deleted text is discarded":"will be removed permanently"):item[0]==="Personal metadata"?"found "+count+" values ("+presentPersonal.join(", ")+") · review or edit above":"found "+count+" · not removed by this tool"):"none found";
      li.textContent=item[0]+": "+detail;
      box.appendChild(li);
    });
  }
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
    add("Comments property",orig.desc,$("desc").value);
    add("Created",fmtD(orig.created),fmtD($("created").value));
    add("Last modified",fmtD(orig.modified),fmtD($("modified").value));
    add("Editing time",fmtT(parseInt(orig.minutes,10)||0),fmtT(getTime()));
    var ids=selectedRules();
    if(ids.length){
      var labels=RULES.concat(CHARS).filter(function(rule){return ids.indexOf(rule.id)>-1}).map(function(rule){return rule.label});
      var matches=ids.reduce(function(sum,id){return sum+(ruleCounts[id]||0)},0);
      rows.push(["Text cleanups",labels.join(", "),matches+" match"+(matches===1?"":"es")+" estimated"]);
    }
    if($("accept").checked)rows.push(["Tracked changes",privacyFindings.tracked+" tracked changes","will be accepted"]);
    if($("removeComments").checked)rows.push(["Document comments",privacyFindings.comments+" comment"+(privacyFindings.comments===1?"":"s"),"will be removed permanently"]);
    if($("removeCustom").checked)rows.push(["Custom properties",privacyFindings.custom+" propert"+(privacyFindings.custom===1?"y":"ies"),"will be removed permanently"]);
    var box=$("diff");box.textContent="";
    if(!rows.length){box.textContent="Nothing changed yet."}
    else{var t=document.createElement("table"),caption=t.createCaption();caption.textContent="Before and after";var head=t.createTHead().insertRow();
      ["Field","Original","Saved copy"].forEach(function(value){var th=document.createElement("th");th.scope="col";th.textContent=value;head.appendChild(th)});
      var body=t.createTBody();
      rows.forEach(function(r){var tr=body.insertRow();r.forEach(function(v){tr.insertCell().textContent=v})});
      box.appendChild(t)}
    var c=new Date($("created").value),m=new Date($("modified").value);
    $("warn").textContent=(!isNaN(c)&&!isNaN(m)&&m<c)?"Heads up: the last modified date is earlier than the created date.":"";
    renderReview();
  }
  $("editor").addEventListener("input",function(e){
    var current=formSnapshot();
    if(lastState&&!sameSnapshot(current,lastState)){undoHistory.push(lastState);if(undoHistory.length>40)undoHistory.shift()}
    lastState=current;updateUndoButton();renderDiff();if(e.target.closest("#rules,#chars"))saveOpts();
  });
  $("editor").addEventListener("change",function(){setTimeout(function(){lastState=formSnapshot();renderDiff()},0)});
  $("editor").addEventListener("click",function(){setTimeout(function(){renderDiff();saveOpts()},0)});
  $("clearPeople").addEventListener("click",function(){rememberUndo();["author","lastBy","company","subject","keywords","desc"].forEach(function(k){$(k).value=""});renderDiff()});
  function setAll(v){rememberUndo();document.querySelectorAll("#rules input").forEach(function(i){i.checked=v});renderDiff()}
  $("allRules").addEventListener("click",function(){setAll(true)});
  $("noRules").addEventListener("click",function(){setAll(false)});
  $("file").addEventListener("change",function(e){loadFiles(e.target.files)});
  var drop=$("drop");
  ["dragenter","dragover"].forEach(function(t){drop.addEventListener(t,function(e){e.preventDefault();drop.classList.add("over")})});
  ["dragleave","drop"].forEach(function(t){drop.addEventListener(t,function(e){e.preventDefault();drop.classList.remove("over")})});
  drop.addEventListener("drop",function(e){loadFiles(e.dataTransfer.files)});
  $("preview").addEventListener("click",function(){
    var ids=selectedRules(),out=[],box=$("prev");
    box.textContent="";
    if(!zip||!ids.length){box.textContent="Select at least one cleanup to preview.";return}
    mapRuns(docXml,function(t){var n=applyRules(t,ids,{});if(n!==t&&out.length<5)out.push([t,n]);return t});
    if(!out.length){box.textContent="No text would change.";return}
    out.forEach(function(p){var d=document.createElement("p");d.className="fmeta";d.dir="auto";d.textContent=p[0].slice(0,120)+"  →  "+p[1].slice(0,120);box.appendChild(d)});
  });
  $("reset").addEventListener("click",function(){rememberUndo();fill(true);saveOpts();msg("Reset to the original values.","")});
  $("undo").addEventListener("click",function(){
    var previous=undoHistory.pop();if(previous)restoreSnapshot(previous);
    msg("Undid the last change.","ok");
  });
  $("sync").addEventListener("change",function(e){if(e.target.checked){$("lastBy").value=$("author").value;syncUndoState()}});
  $("author").addEventListener("input",function(){if($("sync").checked)$("lastBy").value=$("author").value});
  document.querySelectorAll("[data-now]").forEach(function(b){b.addEventListener("click",function(){rememberUndo();$(b.dataset.now).value=toLocalInput(new Date());renderDiff()})});
  $("afterCreated").addEventListener("click",function(){rememberUndo();$("modified").value=$("created").value;renderDiff()});
  document.querySelectorAll("[data-add]").forEach(function(b){b.addEventListener("click",function(){rememberUndo();setTime(getTime()+parseInt(b.dataset.add,10));renderDiff()})});
  $("zero").addEventListener("click",function(){rememberUndo();setTime(0);renderDiff()});
  [$("hours"),$("mins")].forEach(function(i){i.addEventListener("change",function(){setTime(getTime())})});
  $("clearBatch").addEventListener("click",function(){batchFiles=sourceFile?[sourceFile]:[];$("saveAgain").hidden=true;refreshBatch();msg("Batch cleared. The current document remains open.","ok")});
  var saving=false;
  function setProgress(value,text){
    $("progress").hidden=false;$("progressLabel").hidden=false;
    $("progress").value=Math.max(0,Math.min(100,value));
    $("progressLabel").textContent=text;
  }
  function endProgress(){
    $("progress").hidden=true;$("progressLabel").hidden=true;$("progress").value=0;
    $("editor").removeAttribute("aria-busy");
    saving=false;document.querySelectorAll(".actions button").forEach(function(b){b.disabled=false});
  }
  function formValues(){
    return {
      author:$("author").value,lastBy:$("lastBy").value,title:$("title").value,
      company:$("company").value,subject:$("subject").value,keywords:$("keywords").value,
      desc:$("desc").value,created:toIso($("created").value),modified:toIso($("modified").value),
      minutes:getTime(),ids:selectedRules(),accept:$("accept").checked,
      removeComments:$("removeComments").checked,removeCustom:$("removeCustom").checked
    };
  }
  async function removePackageParts(documentZip,settings){
    var files=Object.keys(documentZip.files);
    if(settings.removeComments){
      files.filter(function(name){return /^word\/(?:comments[^/]*|people)\.xml$/i.test(name)}).forEach(function(name){documentZip.remove(name)});
      await Promise.all(files.filter(function(name){return /^word\/.*\.xml$/i.test(name)&&!/^word\/(?:comments[^/]*|people)\.xml$/i.test(name)}).map(function(name){
        var entry=documentZip.file(name);if(!entry)return;
        return entry.async("string").then(function(xml){
          documentZip.file(name,xml.replace(/<w:(?:commentRangeStart|commentRangeEnd|commentReference)\b[^>]*\/>/g,""));
        });
      }));
    }
    if(settings.removeComments||settings.removeCustom){
      var relFiles=Object.keys(documentZip.files).filter(function(name){return /\.rels$/i.test(name)});
      await Promise.all(relFiles.map(function(name){
        var entry=documentZip.file(name);if(!entry)return;
        return entry.async("string").then(function(xml){
          var updated=xml.replace(/<Relationship\b[^>]*\/>/gi,function(tag){
            var type=(tag.match(/\bType=["']([^"']+)["']/i)||[])[1]||"";
            var target=(tag.match(/\bTarget=["']([^"']+)["']/i)||[])[1]||"";
            var removeComments=settings.removeComments&&/(?:comments|people)/i.test(type);
            var removeCustom=settings.removeCustom&&(/custom-properties/i.test(type)||/custom\.xml$/i.test(target));
            return removeComments||removeCustom?"":tag;
          });
          documentZip.file(name,updated);
        });
      }));
      var types=documentZip.file("[Content_Types].xml");
      if(types)documentZip.file("[Content_Types].xml",await types.async("string").then(function(xml){
        return xml.replace(/<Override\b[^>]*\/>/gi,function(tag){
          var part=(tag.match(/\bPartName=["']([^"']+)["']/i)||[])[1]||"";
          var removeComments=settings.removeComments&&/\/word\/(?:comments[^/]*|people)\.xml$/i.test(part);
          var removeCustom=settings.removeCustom&&/\/docProps\/custom\.xml$/i.test(part);
          return removeComments||removeCustom?"":tag;
        });
      }));
    }
    if(settings.removeCustom)documentZip.remove("docProps/custom.xml");
  }
  async function prepareFile(file,settings,onProgress){
    var documentZip=await JSZip.loadAsync(await file.arrayBuffer());
    var coreEntry=documentZip.file("docProps/core.xml");
    if(!coreEntry)throw new Error(file.name+" has no document properties.");
    var core=await coreEntry.async("string");
    core=setTag(core,"dc:creator",settings.author);
    core=setTag(core,"cp:lastModifiedBy",settings.lastBy);
    core=setTag(core,"dc:title",settings.title);
    core=setTag(core,"dc:subject",settings.subject);
    core=setTag(core,"cp:keywords",settings.keywords);
    core=setTag(core,"dc:description",settings.desc);
    var dt=' xsi:type="dcterms:W3CDTF"';
    core=setTag(core,"dcterms:created",settings.created,dt);
    core=setTag(core,"dcterms:modified",settings.modified,dt);
    documentZip.file("docProps/core.xml",core);
    var appEntry=documentZip.file("docProps/app.xml");
    if(appEntry){
      var app=await appEntry.async("string");
      app=setTag(app,"Company",settings.company);
      app=setTag(app,"TotalTime",String(settings.minutes));
      documentZip.file("docProps/app.xml",app);
    }
    var textNames=Object.keys(documentZip.files).filter(function(name){
      return /^word\/(?:document|header|footer|footnotes|endnotes)\d*\.xml$/i.test(name);
    });
    var changed=0;
    for(var i=0;i<textNames.length;i++){
      var entry=documentZip.file(textNames[i]),xml=await entry.async("string");
      if(settings.accept)xml=acceptAll(xml);
      if(settings.ids.length){
        var counts={};
        xml=mapRuns(xml,function(text){return applyRules(text,settings.ids,counts)});
        Object.keys(counts).forEach(function(id){changed+=counts[id]});
      }
      documentZip.file(textNames[i],xml);
    }
    await removePackageParts(documentZip,settings);
    var blob=await documentZip.generateAsync({
      type:"blob",compression:"DEFLATE",
      mimeType:"application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    },function(metadata){if(onProgress)onProgress(metadata.percent||0)});
    return {blob:blob,changed:changed};
  }
  function downloadBlob(blob,name){
    var url=URL.createObjectURL(blob),a=document.createElement("a");
    a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();
    window.setTimeout(function(){URL.revokeObjectURL(url)},1000);
  }
  function editedName(name){return name.replace(/\.docx$/i,"")+"-edited.docx"}
  function beginSave(){
    if(saving)return false;
    saving=true;document.querySelectorAll(".actions button").forEach(function(b){b.disabled=true});
    $("editor").setAttribute("aria-busy","true");
    $("saveAgain").hidden=true;return true;
  }
  async function saveCurrent(){
    if(!sourceFile||!beginSave())return;
    var settings=formValues();
    if(!settings.created||!settings.modified){endProgress();msg("Enter both a created date and a last modified date.","err");return}
    try{
      setProgress(1,"Preparing "+sourceFile.name+"…");
      var result=await prepareFile(sourceFile,settings,function(percent){setProgress(percent,"Preparing "+sourceFile.name+" · "+Math.round(percent)+"%")});
      downloadBlob(result.blob,editedName(sourceFile.name));
      setProgress(100,"Download ready");
      $("saveAgain").hidden=false;
      msg("Saved "+editedName(sourceFile.name)+(settings.ids.length?" ("+result.changed+" text changes).":"."),"ok");
    }catch(err){msg("Couldn’t save the file: "+((err&&err.message)||"unknown error"),"err")}
    finally{endProgress()}
  }
  async function saveBatch(){
    if(batchFiles.length<2||!beginSave())return;
    var settings=formValues();
    if(!settings.created||!settings.modified){endProgress();msg("Enter both a created date and a last modified date.","err");return}
    try{
      var archive=new JSZip(),names={},total=batchFiles.length;
      for(var i=0;i<total;i++){
        var file=batchFiles[i],base=editedName(file.name),name=base,sequence=2;
        while(names[name]){name=base.replace(/\.docx$/i,"-"+sequence+".docx");sequence++}
        names[name]=true;
        setProgress((i/total)*90,"Editing "+(i+1)+" of "+total+": "+file.name);
        var result=await prepareFile(file,settings,function(percent){
          setProgress(((i+percent/100)/total)*90,"Editing "+(i+1)+" of "+total+": "+Math.round(percent)+"%");
        });
        archive.file(name,result.blob);
      }
      setProgress(92,"Creating batch ZIP…");
      var bundle=await archive.generateAsync({type:"blob",compression:"DEFLATE"},function(metadata){
        setProgress(92+(metadata.percent||0)*.08,"Creating batch ZIP · "+Math.round(metadata.percent||0)+"%");
      });
      downloadBlob(bundle,"metaforge-edited-files.zip");
      setProgress(100,"Batch download ready");
      $("saveAgain").hidden=false;
      msg("Saved "+total+" edited documents in metaforge-edited-files.zip.","ok");
    }catch(err){msg("Couldn’t save the batch: "+((err&&err.message)||"unknown error"),"err")}
    finally{endProgress()}
  }
  $("saveAgain").addEventListener("click",function(){if(batchFiles.length>1)saveBatch();else saveCurrent()});
  $("batchSave").addEventListener("click",saveBatch);
  document.addEventListener("keydown",function(e){
    if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==="s"){e.preventDefault();if(zip)$("editor").requestSubmit()}
  });

  $("editor").addEventListener("submit",function(e){e.preventDefault();if(zip)saveCurrent()});
})();
