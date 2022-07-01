/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
importScripts('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/shim.min.js');
importScripts('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js');
postMessage({t:"ready"});

onmessage = function (evt) {
  var v;
  try {
    v = XLSX.read(evt.data.d, {type: evt.data.b, codepage: evt.data.c});
postMessage({t:"xlsx", d:JSON.stringify(v)});
  } catch(e) { postMessage({t:"e",d:e.stack||e}); }
};