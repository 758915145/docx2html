let PizZip = require('pizzip')
module.exports = convertDocX
function convertDocX() {
  this.relations = {};
  this.header = "";
  return this;
}

convertDocX.prototype.relationships = function (data) {
  let _obj = this;

  data.replace(/<Relationship Id="([^\"]+)" Type="([^\"]+)" Target="([^\"]+)" TargetMode="([^\"]+)"\/>/g, function (m, p1, p2, p3, p4) {
    _obj.relations[p1] = p3;
    return "";
  });

  return this;
}

convertDocX.prototype.xml2html = function (data) {
  // REMOVE ALL TAGS
  data = data.replace(/<[^\>]+>/g, "")

  let header = this.header;

  let _obj = this;

  function cleanMS(s) {
    // smart single quotes and apostrophe
    s = s.replace(/[\u2018\u2019\u201A]/g, "\'");

    // smart double quotes
    s = s.replace(/[\u201C\u201D\u201E]/g, "\"");

    // ellipsis
    s = s.replace(/\u2026/g, "...");

    // dashes
    s = s.replace(/[\u2013\u2014]/g, "-");
    return s;
  }
  function removeEmptyTags(str) {
    return str.replace(/<([^\>\s]+) ?[^\>]*>[\n\r\t\s]*<\/\1>/g, "");
  }
  function makeHyperlinks(str) {
    /*<w:hyperlink r:id=\"rId11\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" w:history=\"1\"><w:r w:rsidR=\"00BD4E61\" w:rsidRPr=\"00BD4E61\"><w:rPr><w:rStyle w:val=\"Hyperlink\"/></w:rPr><w:t>https://docs.google.com/forms/d/e/1FAIpQLSdDrytqID_-0qhB47jq-NO2wWcHT427KPiRufA1mygv7ZMP2g/viewform?c=0&amp;w</w:t></w:r></w:hyperlink>*/
    return str.replace(/<w:hyperlink r:id="([^\"]+)".*?<w:t>([^\<]*)<\/w:t><\/w:r><\/w:hyperlink>/g, function (m, p1, p2) {
      return "[" + p2 + "](" + (_obj.relations[p1] ? _obj.relations[p1] : p1) + ")";
    });
  }

  let mode = function mode(arr) {
    let numMapping = {};
    let greatestFreq = 0;
    let mode;
    arr.forEach(function findMode(number) {
      numMapping[number] = (numMapping[number] || 0) + 1;

      if (greatestFreq < numMapping[number]) {
        greatestFreq = numMapping[number];
        mode = number;
      }
    });
    return +mode;
  }

  header = header.replace(/[\n\r]/g, "");
  data = data.replace(/[\n\r]/g, "");

  header = header.replace(/^.*<w:hdr[^\>]+>(.*)<\/w:hdr>/, function (m, p1) { return p1; });
  data = data.replace(/<w:body>/, function (m, p1) { return "<w:body>" + header; });

  // Remove dodgy Microsoft Word characters
  data = cleanMS(data);
  data = makeHyperlinks(data);
  data = data.replace(/<w:tab\/>/g, " ");

  // Find all used font sizes
  tmp = data;
  bits = tmp.split(/<w:sz /);
  sizes = {};
  size = new Array();
  for (let s = 1; s < bits.length; s++) {
    m = (bits[s].match(/w:val=\"([0-9]+)\"/));
    if (m.length == 2) {
      if (!sizes[m[1]]) sizes[m[1]] = 0;
      sizes[m[1]]++;
      size.push(m[1]);
    }
  }
  normal = mode(size);

  // Convert to DOM
  parser = new DOMParser();
  xmlDoc = parser.parseFromString(data, "text/xml");
  console.log(xmlDoc);
  body = xmlDoc.childNodes[0].childNodes[0];
  html = "";
  // Loop over each item
  for (let i = 0; i < body.childNodes.length; i++) {
    let type = "p";
    let out = "";

    let content = body.childNodes[i].innerHTML.replace(/[\n\r\t\s]/g, "");

    if (body.childNodes[i].nodeName == "w:tbl") type = "table";
    if (body.childNodes[i].innerHTML.indexOf("<w:pStyle w:val=\"ListParagraph\"/>") >= 0) type = "li";
    if (body.childNodes[i].innerHTML.indexOf("<w:numPr>") >= 0) type = "li";

    if (type == "p") {
      m = (body.childNodes[i].innerHTML.match(/sz w:val=\"([0-9]+)\"/));
      fontsize = normal;
      if (m && m.length == 2) {
        fontsize = parseInt(m[1]);
        if (fontsize > normal) type = "h2";
      }
      if (type == "p" && fontsize >= normal) {
        if (body.childNodes[i].childNodes.length > 0) {
          h = body.childNodes[i].childNodes[0].innerHTML;
          if (h) {
            if (h.indexOf("<w:b") >= 0) type = "h3";
          }
        }
      }
      if (type == "p" && body.childNodes[i].innerHTML.indexOf(/\u2022/) >= 0) type = "li";
    }
    if (content) {

      if (type == "table") {
        table = body.childNodes[i];
        out += "<table>";
        row = 0;
        for (let t = 0; t < table.childNodes.length; t++) {
          tr = table.childNodes[t];

          if ((tr.nodeName == "w:tr")) {
            row++;
            if (row == 1) out += "<thead>";
            if (row == 2) out += "<tbody>";
            out += "<tr>";
            for (let c = 0; c < tr.childNodes.length; c++) {
              out += (row > 1 ? "\t<td>" : "\t<th>") + (tr.childNodes[c].innerHTML.replace(/<[^a][^\>]+>/g, "")) + (row > 1 ? "</td>" : "</th>");
            }
            out += "</tr>";
            if (row == 1) out += "</thead>";
          }
        }
        out += "</tbody></table>";

      } else {
        // Build output string
        out = "<" + type + ">" + content + "</" + type + ">";
      }

      html += out;
    }
  }

  /*
    // Remove the head and foot
    data = data.replace(/^.*<w:body>/,"").replace(/<\/w:body><\/w:document>.*$/,"");
  
    data = removeEmptyTags(data);
    data = data.replace(/<w:br ?\/>/g,"\n");
    data = data.replace(/<w:b ?\/>/g,"<BOLD>");
    data = data.replace(/<w:tab\/>/g,"\t");
    data = data.replace(/<w:r><w:rPr>\n<\/w:rPr>\t<\/w:r>/g,"\n\nREMOVED A\n\n");
    data = data.replace(/<w:r [^\>]*><w:rPr>\n<\/w:rPr>\t<\/w:r>/g,"\n\nREMOVED B\n\n");
    data = data.replace(/<w:rPr>\n<\/w:rPr>/g,"\n\nREMOVED C\n\n");
    data = data.replace(/<w:pPr>\n\nREMOVED[^\n]*\n\n<\/w:pPr>/g,"\n\nREMOVED D\n\n");
    data = removeEmptyTags(data);
  	
  	
    data = data.replace(/\n/g,"==NEWLINE==");
    data = data.replace(/<w:p.*>(.*)?<\w:p>/g,function(match,p1){ return "<p>"+p1+"</p>"; });
  */
  
  data = data.replace(/\n\nREMOVED[^\n]*\n\n/g, "\n");

  // Convert MarkDown images back to HTML
  //![image](base64 "Title")
  html = html.replace(/!\[[^\]]+\]\(([^\s]+) \"([^\"]+)\"\)/g, function (m, p1, p2) { return '<img src="data:image/png;base64,' + p1 + '" title="' + p2 + '" />'; })
  html = html.replace(/<h2>((<img [^\>]+\/>)+)<\/h2>/g, function (m, p1) { return "<figure>\n\t" + p1 + "\n</figure>"; });

  // Convert MarkDown links back to HTML
  html = html.replace(/\[([^\]]+)\]\(([^\)]+)\)/g, function (m, p1, p2) { return '<a href="' + p2 + '">' + p1 + '</a>'; })

  // Tidy
  html = html.replace(/(p|h[0-9]|table|figure)><li>/g, function (m, p) { return p + "><ul>\n<li>"; }).replace(/<\/li>[\n\t\s]*<(p|h[0-9]|table|figure)/g, function (m, p) { return "</li></ul>\n<" + p; });
  html = html.replace(/<\/([^\>]*)>/g, function (m, p) { return "<\/" + p + ">\n"; });
  html = html.replace(/<li>/g, "\t<li>");
  html = html.replace(/<li>\u2022 ?/g, "<li>");

  html = html.replace(/\\\* ARABIC ([0-9])/g, function (m, p) { return '&#' + (48 + parseInt(p)) + ';'; });
  return html
}
convertDocX.prototype.buffer2html = async function (buffer) {
  let zip = new PizZip(buffer)
  this.header = "";
  // Load the relationships
  if (zip.files["word/_rels/document.xml.rels"]) {
    this.relationships(zip.files["word/_rels/document.xml.rels"].asText());
  }
  // Load the loadHeader
  if (zip.files["word/header1.xml"]) {
    this.header += zip.files["word/header1.xml"].asText();
  }
  // Now process the rest
  return this.xml2html(zip.files["word/document.xml"].asText());
}
