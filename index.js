var JSZip = require('jszip');
var Docxtemplater = require('docxtemplater');
var expressions= require('angular-expressions');
var fs = require('fs');
var path = require('path');

// Load the docx file as a binary
var content = fs.readFileSync(path.resolve(__dirname, 'input.docx'), 'binary');

var angularParser = function(tag) {
  return {
    // This adds support for the . when using angular expressions  
    get: tag === '.' ? function(s){ return s;} : expressions.compile(tag.replace(/’/g, "'"))
  };
};

var zip = new JSZip(content);
var doc = new Docxtemplater()
  .loadZip(zip)
  .setOptions({parser:angularParser})
  .setData({
    "users": [
      {
        "name": "John",
        "num": 0
      },
      {
        "name": "Mary",
        "num": 6
      },
      {
        "name": "Jane",
        "num": 2
      },
      {
        "name": "Sean",
        "num": 3
      }
    ]
  });

try {
  doc.render();
}
catch (error) {
  var e = {
    message: error.message,
    name: error.name,
    stack: error.stack,
    properties: error.properties,
  }
  console.log(JSON.stringify({error: e}));
  // The error thrown here contains additional information when logged with JSON.stringify (it contains a property object).
  throw error;
}

var buf = doc.getZip().generate({type: 'nodebuffer'});

// buf is a nodejs buffer, you can either write it to a file or do anything else with it.
fs.writeFileSync(path.resolve(__dirname, 'output.docx'), buf);