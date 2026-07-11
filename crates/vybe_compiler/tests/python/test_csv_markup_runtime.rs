//! csv, html, xml.etree, urllib.parse runtime.

crate::runtime_case!(
    csv_reader_row,
    "import csv\nimport io\nr = csv.reader(io.StringIO('a,b\\n1,2'))\nprint(next(r))\n",
    "['a', 'b']"
);
crate::runtime_case!(
    csv_reader_data,
    "import csv\nimport io\nr = csv.reader(io.StringIO('a,b\\n1,2'))\nnext(r)\nprint(next(r))\n",
    "['1', '2']"
);
crate::runtime_case!(
    csv_writer_row,
    "import csv\nimport io\nbuf = io.StringIO()\nw = csv.writer(buf)\nw.writerow(['x', 'y'])\nprint(buf.getvalue().strip())\n",
    "x,y"
);
crate::runtime_case!(
    csv_dictreader,
    "import csv\nimport io\nr = csv.DictReader(io.StringIO('a,b\\n1,2'))\nprint(next(r))\n",
    "{'a': '1', 'b': '2'}"
);
crate::runtime_case!(
    csv_dictwriter,
    "import csv\nimport io\nbuf = io.StringIO()\nw = csv.DictWriter(buf, fieldnames=['a', 'b'])\nw.writeheader()\nprint('a' in buf.getvalue())\n",
    "True"
);
crate::runtime_case!(
    html_escape,
    "import html\nprint(html.escape('<div>'))\n",
    "&lt;div&gt;"
);
crate::runtime_case!(
    html_unescape,
    "import html\nprint(html.unescape('&lt;'))\n",
    "<"
);
crate::runtime_case!(
    html_escape_quote,
    "import html\nprint(html.escape('\"'))\n",
    "&quot;"
);
crate::runtime_case!(
    html_escape_amp,
    "import html\nprint(html.escape('&'))\n",
    "&amp;"
);
crate::runtime_case!(
    xml_fromstring_tag,
    "import xml.etree.ElementTree as ET\nroot = ET.fromstring('<root><child/></root>')\nprint(root.tag)\n",
    "root"
);
crate::runtime_case!(
    xml_fromstring_child,
    "import xml.etree.ElementTree as ET\nroot = ET.fromstring('<root><child x=\"1\"/></root>')\nprint(root[0].tag)\n",
    "child"
);
crate::runtime_case!(
    xml_element_attrib,
    "import xml.etree.ElementTree as ET\nroot = ET.fromstring('<root a=\"1\"/>')\nprint(root.attrib['a'])\n",
    "1"
);
crate::runtime_case!(
    xml_element_text,
    "import xml.etree.ElementTree as ET\nroot = ET.fromstring('<root>hi</root>')\nprint(root.text)\n",
    "hi"
);
crate::runtime_case!(
    xml_element_len,
    "import xml.etree.ElementTree as ET\nroot = ET.fromstring('<root><a/><b/></root>')\nprint(len(root))\n",
    "2"
);
crate::runtime_case!(
    xml_tostring,
    "import xml.etree.ElementTree as ET\nroot = ET.fromstring('<root/>')\nprint('<root' in ET.tostring(root, encoding='unicode'))\n",
    "True"
);
crate::runtime_case!(
    urllib_parse_quote,
    "import urllib.parse\nprint(urllib.parse.quote('a b'))\n",
    "a%20b"
);
crate::runtime_case!(
    urllib_parse_unquote,
    "import urllib.parse\nprint(urllib.parse.unquote('a%20b'))\n",
    "a b"
);
crate::runtime_case!(
    urllib_parse_urlencode,
    "import urllib.parse\nprint(urllib.parse.urlencode({'a': 1, 'b': 2}))\n",
    "a=1&b=2"
);
crate::runtime_case!(
    urllib_parse_urlparse,
    "import urllib.parse\np = urllib.parse.urlparse('http://example.com/path')\nprint(p.netloc)\n",
    "example.com"
);
crate::runtime_case!(
    urllib_parse_parse_qs,
    "import urllib.parse\nprint(urllib.parse.parse_qs('a=1&a=2')['a'])\n",
    "['1', '2']"
);
crate::runtime_case!(
    urllib_parse_urljoin,
    "import urllib.parse\nprint(urllib.parse.urljoin('http://a.com/b/', 'c'))\n",
    "http://a.com/c"
);
crate::runtime_case!(
    urllib_parse_quote_plus,
    "import urllib.parse\nprint(urllib.parse.quote_plus('a b'))\n",
    "a+b"
);
crate::runtime_case!(
    csv_sniffer,
    "import csv\nprint(hasattr(csv, 'Sniffer'))\n",
    "True"
);
crate::runtime_case!(
    csv_quote_minimal,
    "import csv\nprint(csv.QUOTE_MINIMAL)\n",
    "0"
);
crate::runtime_case!(
    html_parser,
    "import html.parser\nprint(hasattr(html.parser, 'HTMLParser'))\n",
    "True"
);
crate::runtime_case!(
    xml_element_find,
    "import xml.etree.ElementTree as ET\nroot = ET.fromstring('<root><item id=\"1\"/></root>')\nprint(root.find('item').get('id'))\n",
    "1"
);
crate::runtime_case!(
    xml_element_iter,
    "import xml.etree.ElementTree as ET\nroot = ET.fromstring('<root><a/><b/></root>')\nprint(len(list(root.iter())))\n",
    "3"
);
crate::runtime_case!(
    urllib_parse_splitquery,
    "import urllib.parse\nprint(urllib.parse.splitquery('path?q=1'))\n",
    "('path', '?q=1')"
);
crate::runtime_case!(
    urllib_parse_splittag,
    "import urllib.parse\nprint(urllib.parse.splittag('path#frag'))\n",
    "('path', '#frag')"
);
crate::runtime_case!(
    urllib_parse_defrag,
    "import urllib.parse\nprint(urllib.parse.defrag('http://x#y')[0])\n",
    "http://x"
);
crate::runtime_case!(
    csv_reader_dialect,
    "import csv\nprint(csv.excel.delimiter)\n",
    ","
);
crate::runtime_case!(
    html_entities_map,
    "import html\nprint('&lt;' in html.entities.html5)\n",
    "True"
);
crate::runtime_case!(
    xml_comment,
    "import xml.etree.ElementTree as ET\nroot = ET.fromstring('<root><!--c--><a/></root>')\nprint(len(root))\n",
    "1"
);
crate::runtime_case!(
    urllib_parse_unquote_to_bytes,
    "import urllib.parse\nprint(urllib.parse.unquote_to_bytes('a%20b'))\n",
    "b'a b'"
);
crate::runtime_case!(
    urllib_parse_urldefrag,
    "import urllib.parse\nprint(urllib.parse.urldefrag('http://x#f')[1])\n",
    "#f"
);
crate::runtime_case!(
    csv_list_dialects,
    "import csv\nprint(isinstance(csv.list_dialects(), list))\n",
    "True"
);
crate::runtime_case!(
    html_escape_single,
    "import html\nprint(html.escape(\"'\"))\n",
    "&#x27;"
);
crate::runtime_case!(
    xml_element_makeelement,
    "import xml.etree.ElementTree as ET\nroot = ET.Element('root')\nchild = ET.SubElement(root, 'child')\nprint(child.tag)\n",
    "child"
);
crate::runtime_case!(
    urllib_parse_uses_netloc,
    "import urllib.parse\nprint(isinstance(urllib.parse.uses_netloc, list))\n",
    "True"
);
crate::runtime_case!(
    csv_field_size_limit,
    "import csv\nprint(csv.field_size_limit() > 0)\n",
    "True"
);
crate::runtime_case!(
    xml_parse_error,
    "import xml.etree.ElementTree as ET\ntry:\n ET.fromstring('<unclosed')\n print('ok')\nexcept ET.ParseError:\n print('err')\n",
    "err"
);
crate::runtime_case!(
    urllib_parse_quote_from_bytes,
    "import urllib.parse\nprint(urllib.parse.quote_from_bytes(b'a b'))\n",
    "a%20b"
);
crate::runtime_case!(
    html_unescape_numeric,
    "import html\nprint(html.unescape('&#65;'))\n",
    "A"
);
crate::runtime_case!(
    xml_element_clear,
    "import xml.etree.ElementTree as ET\nroot = ET.fromstring('<root><a/></root>')\nroot.clear()\nprint(len(root))\n",
    "0"
);
crate::runtime_case!(
    urllib_parse_parse_qsl,
    "import urllib.parse\nprint(urllib.parse.parse_qsl('a=1&b=2'))\n",
    "[('a', '1'), ('b', '2')]"
);

crate::compile_case!(
    csv_register_dialect,
    "import csv\ncsv.register_dialect('x', delimiter=';')\n"
);
crate::compile_case!(
    html_parser_feed,
    "from html.parser import HTMLParser\np = HTMLParser()\np.feed('<p>')\n"
);
crate::compile_case!(
    xml_iterparse,
    "import xml.etree.ElementTree as ET\nET.iterparse\n"
);
crate::compile_case!(
    urllib_request,
    "import urllib.request\nurllib.request.urlopen\n"
);
crate::compile_case!(urllib_error, "import urllib.error\nurllib.error.URLError\n");
