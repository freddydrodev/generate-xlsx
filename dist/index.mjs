var H=Object.defineProperty,M=Object.defineProperties;var v=Object.getOwnPropertyDescriptors;var P=Object.getOwnPropertySymbols;var I=Object.prototype.hasOwnProperty,X=Object.prototype.propertyIsEnumerable;var U=(i,r,e)=>r in i?H(i,r,{enumerable:!0,configurable:!0,writable:!0,value:e}):i[r]=e,k=(i,r)=>{for(var e in r||(r={}))I.call(r,e)&&U(i,e,r[e]);if(P)for(var e of P(r))X.call(r,e)&&U(i,e,r[e]);return i},B=(i,r)=>M(i,v(r));var S=(i,r,e)=>new Promise((h,E)=>{var c=l=>{try{a(e.next(l))}catch(s){E(s)}},d=l=>{try{a(e.throw(l))}catch(s){E(s)}},a=l=>l.done?h(l.value):Promise.resolve(l.value).then(c,d);a((e=e.apply(i,r)).next())});import $ from"exceljs";import{saveAs as j}from"file-saver";var u={bottom:{color:{argb:"FF000000"},style:"thin"},top:{color:{argb:"FF000000"},style:"thin"},left:{color:{argb:"FF000000"},style:"thin"},right:{color:{argb:"FF000000"},style:"thin"}},y={horizontal:"center",vertical:"middle",wrapText:!0},b=50,x={size:16},A=B(k({},x),{bold:!0}),W="#\u202F## [$F\u202FCFA-fr-CI];-#\u202F## [$F\u202FCFA-fr-CI]";var Z=i=>S(void 0,null,function*(){var R,D,L,N,O;let{data:r,config:e,rowAlignment:h,defaultFont:E,height:c,fileName:d,headers:a}=i,l=new $.Workbook,s=l.addWorksheet(e.name,{views:[{style:"pageBreakPreview"}],properties:{defaultRowHeight:(R=e.colHeight)!=null?R:b},headerFooter:{oddFooter:"&F&RPage &P / &N"},pageSetup:{paperSize:9,horizontalCentered:!0,scale:(D=e.zoom)!=null?D:100,orientation:(L=e.orientation)!=null?L:"portrait",margins:{top:.75,bottom:.75,left:.25,right:.25,header:.3,footer:.3}}});s.columns=a,s.eachRow(o=>{o.height=c!=null?c:b}),s.columns.forEach((o,m)=>{var n;(n=o.eachCell)==null||n.call(o,{includeEmpty:!1},g=>{var f,F,p,C;let t=a.at(m);g.border=(f=t==null?void 0:t.border)!=null?f:u,g.font=(F=t==null?void 0:t.font)!=null?F:A,o.alignment=(p=t==null?void 0:t.alignment)!=null?p:y,g.fill={pattern:"solid",type:"pattern",fgColor:{argb:"FFD9D9D9"}},(t!=null&&t.isCurrency||t!=null&&t.isNumber)&&(o.numFmt=((C=t==null?void 0:t.numFmt)!=null?C:t.isCurrency)?W:"#,##;-#,##")})});let _=s.addRows(r);_.length>0&&_.forEach((o,m)=>{o.height=b,o.eachCell({includeEmpty:!1},(n,g)=>{var f,F,p;let t=a.at(g-1);n.border=(f=t==null?void 0:t.border)!=null?f:u,n.alignment=(F=t==null?void 0:t.alignment)!=null?F:y,n.font=(p=t==null?void 0:t.font)!=null?p:x})});let T={};for(let o=0;o<a.length;o++){let m=a[o];m.hasTotal?T[(N=m.key)!=null?N:"-"]=0:T[(O=m.key)!=null?O:"-"]=null}let w=s.addRow(T);if(w.height=b,w.eachCell({includeEmpty:!0},(o,m)=>{var g,t,f;let n=a[m-1];if(o.border=(g=n==null?void 0:n.border)!=null?g:u,o.alignment=(t=n==null?void 0:n.alignment)!=null?t:y,o.font=(f=n==null?void 0:n.font)!=null?f:A,n!=null&&n.hasTotal){console.log(n);let F=o.address.replace(/[0-9]+/gi,"");o.value={date1904:!1,formula:`SUM(${F+"1"}:${F+(o.row-1)})`}}}),r.length<=0||a.length<=0)return;let z=yield l.xlsx.writeBuffer(),G=new Blob([z],{type:"applicationi/xlsx"});j(G,d.replace(/\.xlsx/gi,"")+".xlsx")});export{Z as generateXLSXGrid};
//# sourceMappingURL=index.mjs.map