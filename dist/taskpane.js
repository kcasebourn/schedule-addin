!function(){"use strict";var e={14183:function(e,t){var n=this&&this.__awaiter||function(e,t,n,r){return new(n||(n=Promise))((function(s,a){function c(e){try{u(r.next(e))}catch(e){a(e)}}function o(e){try{u(r.throw(e))}catch(e){a(e)}}function u(e){var t;e.done?s(e.value):(t=e.value,t instanceof n?t:new n((function(e){e(t)}))).then(c,o)}u((r=r.apply(e,t||[])).next())}))},r=this&&this.__generator||function(e,t){var n,r,s,a,c={label:0,sent:function(){if(1&s[0])throw s[1];return s[1]},trys:[],ops:[]};return a={next:o(0),throw:o(1),return:o(2)},"function"==typeof Symbol&&(a[Symbol.iterator]=function(){return this}),a;function o(o){return function(u){return function(o){if(n)throw new TypeError("Generator is already executing.");for(;a&&(a=0,o[0]&&(c=0)),c;)try{if(n=1,r&&(s=2&o[0]?r.return:o[0]?r.throw||((s=r.return)&&s.call(r),0):r.next)&&!(s=s.call(r,o[1])).done)return s;switch(r=0,s&&(o=[2&o[0],s.value]),o[0]){case 0:case 1:s=o;break;case 4:return c.label++,{value:o[1],done:!1};case 5:c.label++,r=o[1],o=[0];continue;case 7:o=c.ops.pop(),c.trys.pop();continue;default:if(!((s=(s=c.trys).length>0&&s[s.length-1])||6!==o[0]&&2!==o[0])){c=0;continue}if(3===o[0]&&(!s||o[1]>s[0]&&o[1]<s[3])){c.label=o[1];break}if(6===o[0]&&c.label<s[1]){c.label=s[1],s=o;break}if(s&&c.label<s[2]){c.label=s[2],c.ops.push(o);break}s[2]&&c.ops.pop(),c.trys.pop();continue}o=t.call(e,c)}catch(e){o=[6,e],r=0}finally{n=s=0}if(5&o[0])throw o[1];return{value:o[0]?o[1]:void 0,done:!0}}([o,u])}}};function s(){return n(this,void 0,void 0,(function(){var e,t=this;return r(this,(function(s){switch(s.label){case 0:return s.trys.push([0,2,,3]),[4,Excel.run((function(e){return n(t,void 0,void 0,(function(){var t;return r(this,(function(n){switch(n.label){case 0:return(t=e.workbook.getSelectedRange()).load("address"),[4,e.sync()];case 1:return n.sent(),console.log("The range address was ".concat(t.address,".")),[2]}}))}))}))];case 1:return s.sent(),[3,3];case 2:return e=s.sent(),console.error(e),[3,3];case 3:return[2]}}))}))}function a(e){return n(this,void 0,void 0,(function(){var t=this;return r(this,(function(s){switch(s.label){case 0:return[4,Excel.run({delayForCellEdit:!0},(function(s){return n(t,void 0,void 0,(function(){var t,n,a;return r(this,(function(r){switch(r.label){case 0:return"I"!==e.address.substring(0,1)||"ThisLocalAddin"===e.triggerSource?[2]:(t=s.workbook.worksheets.getItem("Schedule"),n=t.getRange("I103:I200"),a=[{key:0,ascending:!0}],n.sort.apply(a),[4,s.sync()]);case 1:return r.sent(),[2]}}))}))}))];case 1:return s.sent(),[2]}}))}))}function c(e){return n(this,void 0,void 0,(function(){var t=this;return r(this,(function(s){switch(s.label){case 0:return[4,Excel.run({delayForCellEdit:!0},(function(s){return n(t,void 0,void 0,(function(){var t,a,c,o,u,i,g;return r(this,(function(d){switch(d.label){case 0:if(d.trys.push([0,7,,8]),"ThisLocalAddin"===e.triggerSource)return[2];if(e.address.includes(":"))return function(){n(this,void 0,void 0,(function(){var e=this;return r(this,(function(t){switch(t.label){case 0:return[4,Excel.run({delayForCellEdit:!0},(function(t){return n(e,void 0,void 0,(function(){var e,n,s;return r(this,(function(r){switch(r.label){case 0:return e=t.workbook.worksheets.getItem("Drivers"),n=e.getRange("A4:P100"),s=[{key:0,ascending:!0}],n.sort.apply(s),[4,t.sync()];case 1:return r.sent(),e.getRange("$M$4").autoFill("$M4:$M100"),[4,t.sync()];case 2:return r.sent(),[2]}}))}))}))];case 1:return t.sent(),[2]}}))}))}(),[2];switch(t=s.workbook.worksheets.getItem("Drivers"),a=t.getRange(e.address),c=e.details,o=c.valueAfter,e.address.substring(0,1)){case"A":case"D":case"J":case"N":return[3,1];case"B":case"O":return[3,2];case"V":return[3,3]}return[3,5];case 1:return a.values=o.toString().toUpperCase(),[3,5];case 2:return a.values=l(o),[3,5];case 3:return a.values=l(o),[4,s.sync()];case 4:return d.sent(),u=t.getRange("A4:V100"),i=[{key:0,ascending:!0}],u.sort.apply(i),[3,5];case 5:return[4,s.sync()];case 6:return d.sent(),[3,8];case 7:return g=d.sent(),console.log(g),[3,8];case 8:return[2]}}))}))}))];case 1:return s.sent(),[2]}}))}))}function o(e){return n(this,void 0,void 0,(function(){var t=this;return r(this,(function(s){switch(s.label){case 0:return[4,Excel.run({delayForCellEdit:!0},(function(s){return n(t,void 0,void 0,(function(){var t,n,a,c,o,u,i,g,f;return r(this,(function(r){switch(r.label){case 0:return r.trys.push([0,7,,8]),t=e.address,"ThisLocalAddin"===e.triggerSource?[2]:(n=s.workbook.worksheets.getItem("Schedule By Start Time"),a=e.details,c="",void 0===a?[3,5]:"STOP"!==a.valueAfter&&"stop"!==a.valueAfter&&"Stop"!==a.valueAfter?[3,3]:((o=n.getRange("A1")).load("values"),[4,s.sync()]));case 1:return r.sent(),u=o.values,clearInterval(u[0][0].valueOf()),[4,s.sync()];case 2:return r.sent(),[2];case 3:c=t.substring(0,1),i=a.valueAfter,r.label=4;case 4:return[3,6];case 5:c="E",r.label=6;case 6:switch(g=n.getRange(t),c){case"B":g.values=l(i);break;case"D":case"F":g.values=i.toString().toUpperCase();break;case"E":d()}return s.sync(),[3,8];case 7:return f=r.sent(),console.error(f),[3,8];case 8:return[2]}}))}))}))];case 1:return s.sent(),[2]}}))}))}function u(e){return n(this,void 0,void 0,(function(){var t=this;return r(this,(function(s){switch(s.label){case 0:return[4,Excel.run({delayForCellEdit:!0},(function(s){return n(t,void 0,void 0,(function(){var t,n,a,c,o,u,i,l;return r(this,(function(r){switch(r.label){case 0:t=e.address,n=s.workbook.worksheets.getActiveWorksheet(),r.label=1;case 1:return r.trys.push([1,4,,6]),n.protection.pauseProtection("CPM"),(a=n.getRange(t)).load("values"),[4,s.sync()];case 2:switch(r.sent(),c=a.values,o=c.toString(),t){case"AB301":case"D301":"-"===o?(n.getRange("D:AA").columnHidden=!0,(u=n.getRange("AB301")).values="+",u.getOffsetRange(1,0).select()):(n.getRange("D:AA").columnHidden=!1,(i=n.getRange("AB301")).values="-",i.getOffsetRange(1,0).select());break;case"AC301":case"BA301":"-"===o?(n.getRange("AC:AZ").columnHidden=!0,(u=n.getRange("BA301")).values="+",u.getOffsetRange(1,0).select()):(n.getRange("AC:AZ").columnHidden=!1,(i=n.getRange("BA301")).values="-",i.getOffsetRange(1,0).select());break;case"BB301":case"BZ301":"-"===o?(n.getRange("BB:BY").columnHidden=!0,(u=n.getRange("BZ301")).values="+",u.getOffsetRange(1,0).select()):(n.getRange("BB:BY").columnHidden=!1,(i=n.getRange("BZ301")).values="-",i.getOffsetRange(1,0).select());break;case"CA301":case"CY301":"-"===o?(n.getRange("CA:CX").columnHidden=!0,(u=n.getRange("CY301")).values="+",u.getOffsetRange(1,0).select()):(n.getRange("CA:CX").columnHidden=!1,(i=n.getRange("CY301")).values="-",i.getOffsetRange(1,0).select());break;case"CZ301":case"DX301":"-"===o?(n.getRange("CZ:DW").columnHidden=!0,(u=n.getRange("DX301")).values="+",u.getOffsetRange(1,0).select()):(n.getRange("CZ:DW").columnHidden=!1,(i=n.getRange("DX301")).values="-",i.getOffsetRange(1,0).select());break;case"DY301":case"EW301":"-"===o?(n.getRange("DY:EV").columnHidden=!0,(u=n.getRange("EW301")).values="+",u.getOffsetRange(1,0).select()):(n.getRange("DY:EV").columnHidden=!1,(i=n.getRange("EW301")).values="-",i.getOffsetRange(1,0).select());break;case"EX301":case"FV301":"-"===o?(n.getRange("EX:FU").columnHidden=!0,(u=n.getRange("FV301")).values="+",u.getOffsetRange(1,0).select()):(n.getRange("EX:FU").columnHidden=!1,(i=n.getRange("FV301")).values="-",i.getOffsetRange(1,0).select())}return[4,s.sync()];case 3:return r.sent(),n.protection.resumeProtection(),[3,6];case 4:return l=r.sent(),console.error(l),[4,s.sync()];case 5:return r.sent(),n.protection.resumeProtection(),[3,6];case 6:return[2]}}))}))}))];case 1:return s.sent(),[2]}}))}))}function i(e){return n(this,void 0,void 0,(function(){var t=this;return r(this,(function(s){switch(s.label){case 0:return[4,Excel.run({delayForCellEdit:!0},(function(s){return n(t,void 0,void 0,(function(){var t,n,a,c,o,u,i,l;return r(this,(function(r){switch(r.label){case 0:t=e.address,n=s.workbook.worksheets.getActiveWorksheet(),r.label=1;case 1:return r.trys.push([1,4,,6]),n.protection.pauseProtection("CPM"),(a=n.getRange(t)).load("values"),[4,s.sync()];case 2:switch(r.sent(),c=a.values,o=c.toString(),t){case"AE201":case"G201":"-"===o?(n.getRange("G:AD").columnHidden=!0,(u=n.getRange("AE201")).values="+",u.getOffsetRange(1,0).select()):(n.getRange("G:AD").columnHidden=!1,(i=n.getRange("AE201")).values="-",i.getOffsetRange(1,0).select());break;case"AF201":case"BD201":"-"===o?(n.getRange("AF:BC").columnHidden=!0,(u=n.getRange("BD201")).values="+",u.getOffsetRange(1,0).select()):(n.getRange("AF:BC").columnHidden=!1,(i=n.getRange("BD201")).values="-",i.getOffsetRange(1,0).select());break;case"BE201":case"CC201":"-"===o?(n.getRange("BE:CB").columnHidden=!0,(u=n.getRange("CC201")).values="+",u.getOffsetRange(1,0).select()):(n.getRange("BE:CB").columnHidden=!1,(i=n.getRange("CC201")).values="-",i.getOffsetRange(1,0).select());break;case"CD201":case"DB201":"-"===o?(n.getRange("CD:DA").columnHidden=!0,(u=n.getRange("DB201")).values="+",u.getOffsetRange(1,0).select()):(n.getRange("CD:DA").columnHidden=!1,(i=n.getRange("DB201")).values="-",i.getOffsetRange(1,0).select());break;case"DC201":case"EA201":"-"===o?(n.getRange("DC:DZ").columnHidden=!0,(u=n.getRange("EA201")).values="+",u.getOffsetRange(1,0).select()):(n.getRange("DC:DZ").columnHidden=!1,(i=n.getRange("EA201")).values="-",i.getOffsetRange(1,0).select());break;case"EB201":case"EZ201":"-"===o?(n.getRange("EB:EY").columnHidden=!0,(u=n.getRange("EZ201")).values="+",u.getOffsetRange(1,0).select()):(n.getRange("EB:EY").columnHidden=!1,(i=n.getRange("EZ201")).values="-",i.getOffsetRange(1,0).select());break;case"FA201":case"FY201":"-"===o?(n.getRange("FA:FX").columnHidden=!0,(u=n.getRange("FY201")).values="+",u.getOffsetRange(1,0).select()):(n.getRange("FA:FX").columnHidden=!1,(i=n.getRange("FY201")).values="-",i.getOffsetRange(1,0).select())}return[4,s.sync()];case 3:return r.sent(),n.protection.resumeProtection(),[3,6];case 4:return l=r.sent(),console.error(l),[4,s.sync()];case 5:return r.sent(),n.protection.resumeProtection(),[3,6];case 6:return[2]}}))}))}))];case 1:return s.sent(),[2]}}))}))}function l(e){return String(e).toLowerCase().replace(/(^| )(\w)/g,(function(e){return e.toUpperCase()}))}function g(){return n(this,void 0,void 0,(function(){var e=this;return r(this,(function(t){switch(t.label){case 0:return[4,Excel.run({delayForCellEdit:!0},(function(t){return n(e,void 0,void 0,(function(){var e,n,s,a;return r(this,(function(r){switch(r.label){case 0:return e=t.workbook.worksheets.getItem("Schedule By Start Time"),d(),(n=e.getRange("F1")).load("values"),[4,t.sync()];case 1:return r.sent(),s=n.values,a=+s+5,n.values=a,[4,t.sync()];case 2:return r.sent(),[2]}}))}))}))];case 1:return t.sent(),[2]}}))}))}function d(){return n(this,void 0,void 0,(function(){var e=this;return r(this,(function(t){return Excel.run({delayForCellEdit:!0},(function(t){return n(e,void 0,void 0,(function(){var e,n,s,a,c,o,u,i,l,g,d,f,v,h,R,b,p,y,w,m,k,C,E,A,B,O,D;return r(this,(function(r){switch(r.label){case 0:return r.trys.push([0,4,,5]),e=t.workbook.worksheets.getItem("Schedule By Start Time"),n=e.getRange("AW2"),s=e.getRange("C101"),n.load("values"),s.load("values"),[4,t.sync()];case 1:return r.sent(),a=n.values,103==+(c=s.values)&&(c="104"),o=e.getRange("AW3:BA"+a),u=e.getRange("B104:F"+c),o.load("values"),u.load("values"),[4,t.sync()];case 2:for(r.sent(),i=o.values,l=u.values,g=l.length,d=[],d=""!==l[0][0]?l.concat(i):i,f=d.length,v=[],h=!1,R=new Date,b=R.getMinutes()+"",p=R.getHours()+"",b.length<2&&(b="0"+b),p.length<2&&(p="0"+p),y=p+b,O=0;O<f;O++)if(""!==d[O][0]||""!==d[O][1]||""!==d[O][2]||""!==d[O][3]||""!==d[O][4]){if((w=+d[O][1]+1200)<2400){if(+y>w&&+y>+d[O][3]&&""!==d[O][1])continue}else if(0,m=+y<1200?+y+2400:+y,0,k=+d[O][3]<1200?+d[O][3]+2400:+d[O][3],m>w&&m>k&&""!==d[O][1])continue;for(C=d[O][0].toString().toLowerCase(),E=v.length,A=0;A<E;A++)if(C===v[A][0].toString().toLowerCase()){h=!0;break}h?h=!1:v.push([d[O][0],d[O][1],d[O][2],d[O][3],d[O][4]])}if(v.sort((function(e,t){var n=e[4],r=t[4],s=n==r?0:n<r?-1:1;if(0==s){var a=e[3],c=t[3];s=""===a&&""!==c?1:""!==a&&""===c||+a-+c>1200?-1:+c-+a>1200?1:a==c?0:a<c?-1:1}if(0==s){var o=e[1],u=t[1];s=o==u?0:o<u?-1:1}return s})),B=v.length,g>B)for(O=B;O<g;O++)v.push(["","","","",""]);return e.getRange("B104:F"+(103+v.length)).values=v,[4,t.sync()];case 3:return r.sent(),[3,5];case 4:return D=r.sent(),console.error(D),[3,5];case 5:return[2]}}))}))})),[2]}))}))}Object.defineProperty(t,"__esModule",{value:!0}),t.run=void 0,Office.onReady().then((function(){var e=this;document.getElementById("sideload-msg").style.display="none",document.getElementById("app-body").style.display="flex",document.getElementById("run").onclick=s,Office.addin.setStartupBehavior(Office.StartupBehavior.load),Excel.run((function(t){return n(e,void 0,void 0,(function(){var e,s,l,f,v,h,R,b;return r(this,(function(p){switch(p.label){case 0:return function(){n(this,void 0,void 0,(function(){var e=this;return r(this,(function(t){return Excel.run((function(t){return n(e,void 0,void 0,(function(){return r(this,(function(e){switch(e.label){case 0:return t.runtime.load("enableEvents"),[4,t.sync()];case 1:return e.sent(),t.runtime.enableEvents=!0,[2]}}))}))})),[2]}))}))}(),e=t.workbook.worksheets.getItem("Schedule"),s=t.workbook.worksheets.getItem("Drivers"),l=t.workbook.worksheets.getItem("Schedule By Start Time"),f=t.workbook.worksheets.getItem("TUC"),v=t.workbook.worksheets.getItem("DUC"),d(),h=l.getRange("A1"),R=l.getRange("F1"),b=setInterval(g,3e5),R.values=0,h.values=b,[4,t.sync()];case 1:return p.sent(),e.onChanged.add(a),s.onChanged.add(c),l.onChanged.add(o),f.onSelectionChanged.add(u),v.onSelectionChanged.add(i),[4,t.sync()];case 2:return p.sent(),[2]}}))}))}))})),t.run=s},93823:function(e,t,n){var r=n(27091),s=n.n(r),a=new URL(n(60806),n.b);s()(a)},27091:function(e){e.exports=function(e,t){return t||(t={}),e?(e=String(e.__esModule?e.default:e),t.hash&&(e+=t.hash),t.maybeNeedQuotes&&/[\t\n\f\r "'=<>`]/.test(e)?'"'.concat(e,'"'):e):e}},60806:function(e,t,n){e.exports=n.p+"7791e6ec138c3048b5f3.css"}},t={};function n(r){var s=t[r];if(void 0!==s)return s.exports;var a=t[r]={exports:{}};return e[r].call(a.exports,a,a.exports,n),a.exports}n.m=e,n.n=function(e){var t=e&&e.__esModule?function(){return e.default}:function(){return e};return n.d(t,{a:t}),t},n.d=function(e,t){for(var r in t)n.o(t,r)&&!n.o(e,r)&&Object.defineProperty(e,r,{enumerable:!0,get:t[r]})},n.g=function(){if("object"==typeof globalThis)return globalThis;try{return this||new Function("return this")()}catch(e){if("object"==typeof window)return window}}(),n.o=function(e,t){return Object.prototype.hasOwnProperty.call(e,t)},function(){var e;n.g.importScripts&&(e=n.g.location+"");var t=n.g.document;if(!e&&t&&(t.currentScript&&(e=t.currentScript.src),!e)){var r=t.getElementsByTagName("script");r.length&&(e=r[r.length-1].src)}if(!e)throw new Error("Automatic publicPath is not supported in this browser");e=e.replace(/#.*$/,"").replace(/\?.*$/,"").replace(/\/[^\/]+$/,"/"),n.p=e}(),n.b=document.baseURI||self.location.href,n(14183),n(93823)}();
//# sourceMappingURL=taskpane.js.map