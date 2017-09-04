var X = XLSX;
var process_wb = (function() {
  var OUT = document.getElementById('out');
  var HTMLOUT = document.getElementById('htmlout');
  var get_format = (function() {
    var radios = document.getElementsByName( "format" );
    return function() {
      for(var i = 0; i < radios.length; ++i) if(radios[i].checked || radios.length === 1) return radios[i].value;
    };
  })();
  var to_json = function to_json(workbook) {
    var result = {};
    workbook.SheetNames.forEach(function(sheetName) {
      var roa = X.utils.sheet_to_json(workbook.Sheets[sheetName]);
      if(roa.length) result[sheetName] = roa;
    });
    // return JSON.stringify(result, 2, 2);
    return result;
  };
  var to_csv = function to_csv(workbook) {
    var result = [];
    workbook.SheetNames.forEach(function(sheetName) {
      var csv = X.utils.sheet_to_csv(workbook.Sheets[sheetName]);
      if(csv.length){
        result.push("SHEET: " + sheetName);
        result.push("");
        result.push(csv);
      }
    });
    return result.join("\n");
  };
  var to_fmla = function to_fmla(workbook) {
    var result = [];
    workbook.SheetNames.forEach(function(sheetName) {
      var formulae = X.utils.get_formulae(workbook.Sheets[sheetName]);
      if(formulae.length){
        result.push("SHEET: " + sheetName);
        result.push("");
        result.push(formulae.join("\n"));
      }
    });
    return result.join("\n");
  };
  var to_html = function to_html(workbook) {
    HTMLOUT.innerHTML = "";
    workbook.SheetNames.forEach(function(sheetName) {
      var htmlstr = X.write(workbook, {sheet:sheetName, type:'binary', bookType:'html'});
      HTMLOUT.innerHTML += htmlstr;
    });
    return "";
  };
  return function process_wb(wb, cb) {
    console.log("process_wb");
    global_wb = wb;
    var output = "";
    // switch(get_format()) {
    //   case "form": output = to_fmla(wb); break;
    //   case "html": output = to_html(wb); break;
    //   case "json": output = to_json(wb); break;
    //   default: output = to_csv(wb);
    // }
    output = to_json(wb);


    // if(OUT.innerText === undefined) OUT.textContent = output;
    // else OUT.innerText = output;
    // if(typeof console !== 'undefined') console.log("output", new Date());
    cb(output);
  };
})();



var do_file = (function(cb) {
  console.log("dofile");
  var rABS = typeof FileReader !== "undefined" && (FileReader.prototype||{}).readAsBinaryString;
  // var domrabs = document.getElementsByName("userabs")[0];
  // if(!rABS) domrabs.disabled = !(domrabs.checked = false);
  // var use_worker = typeof Worker !== 'undefined';
  // var domwork = document.getElementsByName("useworker")[0];
  // if(!use_worker) domwork.disabled = !(domwork.checked = false);
  var xw = function xw(data, cb) {
    var worker = new Worker(XW.worker);
    worker.onmessage = function(e) {
      switch(e.data.t) {
        case 'ready': break;
        case 'e': console.error(e.data.d); break;
        case XW.msg: cb(JSON.parse(e.data.d)); break;
      }
    };
    worker.postMessage({d:data,b:rABS?'binary':'array'});
  };
  return function do_file(files, cb) {
    // rABS = domrabs.checked;
    rABS = false;
    // use_worker = domwork.checked;
    use_worker = false;
    // var f = files[0];
    // convert into File from blob
    // var f = new File([files], "conent.xlsx", {type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", lastModified: Date.now()});
    var f = new File([files], "conent.xlsx");
    var reader = new FileReader();
    reader.onload = function(e) {
      if(typeof console !== 'undefined') console.log("onload", new Date(), rABS, use_worker);
      var data = e.target.result;
      console.log("data", e);
      if(!rABS) data = new Uint8Array(data);
      if(use_worker) xw(data, process_wb);
      else process_wb(X.read(data, {type: rABS ? 'binary' : 'array'}), cb);
    };
    if(rABS) reader.readAsBinaryString(f);
    else reader.readAsArrayBuffer(f);
  };
})();

var xmlHttpRequest = new XMLHttpRequest();


var getFileBlob = function (url, cb) {
  var xhr = new XMLHttpRequest();
  xhr.open("GET", url);
  xhr.responseType = "blob";
  xhr.addEventListener('load', function() {
      cb(xhr.response);
  });
  xhr.send();
};

var blobToFile = function (blob, name) {
      blob.lastModifiedDate = new Date();
      blob.name = name;
      return blob;
};

var getFileObject = function(filePathOrUrl, cb) {
     getFileBlob(filePathOrUrl, function (blob) {
        cb(blobToFile(blob, 'content.xlsx'));
     });
};
