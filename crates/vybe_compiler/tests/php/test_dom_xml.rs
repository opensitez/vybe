use super::helpers::compile_ok;

// ── DOMDocument creation ──────────────────────────────────────

#[test]
fn dom_create_document() {
    compile_ok(
        r#"<?php
$doc = new DOMDocument('1.0', 'UTF-8');
echo $doc->version . ':' . $doc->encoding;
"#,
    );
}

#[test]
fn dom_create_element() {
    compile_ok(
        r#"<?php
$doc = new DOMDocument();
$root = $doc->createElement('root');
$doc->appendChild($root);
echo $doc->documentElement->tagName;
"#,
    );
}

#[test]
fn dom_create_text_node() {
    compile_ok(
        r#"<?php
$doc = new DOMDocument();
$root = $doc->createElement('message');
$text = $doc->createTextNode('Hello, World!');
$root->appendChild($text);
$doc->appendChild($root);
echo $doc->documentElement->textContent;
"#,
    );
}

#[test]
fn dom_create_attribute() {
    compile_ok(
        r#"<?php
$doc = new DOMDocument();
$el = $doc->createElement('person');
$el->setAttribute('name', 'Alice');
$el->setAttribute('age', '30');
$doc->appendChild($el);
echo $el->getAttribute('name') . ':' . $el->getAttribute('age');
"#,
    );
}

#[test]
fn dom_nested_elements() {
    compile_ok(
        r#"<?php
$doc = new DOMDocument('1.0', 'UTF-8');
$root = $doc->createElement('catalog');
$doc->appendChild($root);
$book = $doc->createElement('book');
$book->setAttribute('id', '1');
$title = $doc->createElement('title');
$title->appendChild($doc->createTextNode('PHP Manual'));
$book->appendChild($title);
$root->appendChild($book);
echo $root->childNodes->length;
echo ':' . $root->firstChild->getAttribute('id');
"#,
    );
}

// ── DOMDocument load / save ───────────────────────────────────

#[test]
fn dom_load_xml_string() {
    compile_ok(
        r#"<?php
$xml = '<?xml version="1.0"?><root><item id="1">First</item><item id="2">Second</item></root>';
$doc = new DOMDocument();
$doc->loadXML($xml);
$items = $doc->getElementsByTagName('item');
echo $items->length;
echo ':' . $items->item(0)->textContent;
"#,
    );
}

#[test]
fn dom_save_xml() {
    compile_ok(
        r#"<?php
$doc = new DOMDocument('1.0', 'UTF-8');
$doc->formatOutput = true;
$root = $doc->createElement('data');
$root->appendChild($doc->createTextNode('hello'));
$doc->appendChild($root);
$xml = $doc->saveXML();
echo str_contains($xml, '<data>') ? 'has data tag' : 'missing tag';
echo str_contains($xml, 'hello') ? ':has content' : ':missing content';
"#,
    );
}

#[test]
fn dom_load_html() {
    compile_ok(
        r#"<?php
$html = '<html><body><h1>Title</h1><p class="intro">Paragraph</p></body></html>';
$doc = new DOMDocument();
@$doc->loadHTML($html);
$h1 = $doc->getElementsByTagName('h1')->item(0);
$p  = $doc->getElementsByTagName('p')->item(0);
echo $h1->textContent . ':' . $p->getAttribute('class');
"#,
    );
}

#[test]
fn dom_save_html() {
    compile_ok(
        r#"<?php
$doc = new DOMDocument();
@$doc->loadHTML('<html><body><p>Test</p></body></html>');
$html = $doc->saveHTML();
echo str_contains($html, '<p>Test</p>') ? 'ok' : 'fail';
"#,
    );
}

// ── DOMElement methods ────────────────────────────────────────

#[test]
fn dom_get_elements_by_tag() {
    compile_ok(
        r#"<?php
$xml = '<store><book><title>A</title></book><book><title>B</title></book><book><title>C</title></book></store>';
$doc = new DOMDocument();
$doc->loadXML($xml);
$books = $doc->getElementsByTagName('book');
echo $books->length;
"#,
    );
}

#[test]
fn dom_has_attribute() {
    compile_ok(
        r#"<?php
$doc = new DOMDocument();
$doc->loadXML('<el foo="bar" />');
$el = $doc->documentElement;
echo $el->hasAttribute('foo') ? 'has foo' : 'no foo';
echo $el->hasAttribute('baz') ? 'has baz' : ':no baz';
"#,
    );
}

#[test]
fn dom_remove_attribute() {
    compile_ok(
        r#"<?php
$doc = new DOMDocument();
$doc->loadXML('<el id="1" class="x" />');
$el = $doc->documentElement;
$el->removeAttribute('class');
echo $el->hasAttribute('id')    ? 'id ok' : 'id gone';
echo $el->hasAttribute('class') ? ':class still' : ':class removed';
"#,
    );
}

#[test]
fn dom_clone_node() {
    compile_ok(
        r#"<?php
$doc = new DOMDocument();
$doc->loadXML('<list><item>hello</item></list>');
$item = $doc->getElementsByTagName('item')->item(0);
$clone = $item->cloneNode(true);
$doc->documentElement->appendChild($clone);
echo $doc->getElementsByTagName('item')->length;
"#,
    );
}

#[test]
fn dom_remove_child() {
    compile_ok(
        r#"<?php
$doc = new DOMDocument();
$doc->loadXML('<list><a/><b/><c/></list>');
$root = $doc->documentElement;
$b = $doc->getElementsByTagName('b')->item(0);
$root->removeChild($b);
echo $root->childNodes->length;
"#,
    );
}

#[test]
fn dom_insert_before() {
    compile_ok(
        r#"<?php
$doc = new DOMDocument();
$doc->loadXML('<list><b/><c/></list>');
$root = $doc->documentElement;
$a = $doc->createElement('a');
$b = $doc->getElementsByTagName('b')->item(0);
$root->insertBefore($a, $b);
echo $root->firstChild->tagName;
"#,
    );
}

// ── DOMXPath ─────────────────────────────────────────────────

#[test]
fn dom_xpath_query_basic() {
    compile_ok(
        r#"<?php
$xml = '<store><book price="10"><title>Alpha</title></book><book price="25"><title>Beta</title></book></store>';
$doc = new DOMDocument();
$doc->loadXML($xml);
$xpath = new DOMXPath($doc);
$books = $xpath->query('//book');
echo $books->length;
"#,
    );
}

#[test]
fn dom_xpath_attribute_predicate() {
    compile_ok(
        r#"<?php
$xml = '<items><item id="1" active="true"/><item id="2" active="false"/><item id="3" active="true"/></items>';
$doc = new DOMDocument();
$doc->loadXML($xml);
$xpath = new DOMXPath($doc);
$active = $xpath->query('//item[@active="true"]');
echo $active->length;
"#,
    );
}

#[test]
fn dom_xpath_evaluate() {
    compile_ok(
        r#"<?php
$xml = '<data><val>10</val><val>20</val><val>30</val></data>';
$doc = new DOMDocument();
$doc->loadXML($xml);
$xpath = new DOMXPath($doc);
$count = $xpath->evaluate('count(//val)');
echo (int)$count;
"#,
    );
}

#[test]
fn dom_xpath_text_content() {
    compile_ok(
        r#"<?php
$xml = '<users><user><name>Alice</name><age>30</age></user><user><name>Bob</name><age>25</age></user></users>';
$doc = new DOMDocument();
$doc->loadXML($xml);
$xpath = new DOMXPath($doc);
$names = $xpath->query('//user/name');
$result = [];
foreach ($names as $name) { $result[] = $name->textContent; }
echo implode(',', $result);
"#,
    );
}

#[test]
fn dom_xpath_register_namespace() {
    compile_ok(
        r#"<?php
$xml = '<root xmlns:app="http://example.com/app"><app:item>value</app:item></root>';
$doc = new DOMDocument();
$doc->loadXML($xml);
$xpath = new DOMXPath($doc);
$xpath->registerNamespace('a', 'http://example.com/app');
$items = $xpath->query('//a:item');
echo $items->length . ':' . $items->item(0)->textContent;
"#,
    );
}

// ── XMLWriter ────────────────────────────────────────────────

#[test]
fn xml_writer_basic() {
    compile_ok(
        r#"<?php
$writer = new XMLWriter();
$writer->openMemory();
$writer->startDocument('1.0', 'UTF-8');
$writer->startElement('root');
$writer->writeElement('child', 'value');
$writer->endElement();
$writer->endDocument();
$xml = $writer->outputMemory();
echo str_contains($xml, '<root>') ? 'has root' : 'no root';
echo str_contains($xml, 'value') ? ':has value' : ':no value';
"#,
    );
}

#[test]
fn xml_writer_attributes() {
    compile_ok(
        r#"<?php
$writer = new XMLWriter();
$writer->openMemory();
$writer->startDocument('1.0');
$writer->startElement('person');
$writer->writeAttribute('name', 'Alice');
$writer->writeAttribute('age', '30');
$writer->endElement();
$xml = $writer->outputMemory();
echo str_contains($xml, 'name="Alice"') ? 'has name attr' : 'missing';
echo str_contains($xml, 'age="30"')     ? ':has age attr' : ':missing';
"#,
    );
}

#[test]
fn xml_writer_cdata() {
    compile_ok(
        r#"<?php
$writer = new XMLWriter();
$writer->openMemory();
$writer->startDocument('1.0');
$writer->startElement('code');
$writer->writeCData('<script>alert("xss")</script>');
$writer->endElement();
$xml = $writer->outputMemory();
echo str_contains($xml, 'CDATA') ? 'has CDATA' : 'no CDATA';
"#,
    );
}

#[test]
fn xml_writer_nested() {
    compile_ok(
        r#"<?php
$writer = new XMLWriter();
$writer->openMemory();
$writer->startDocument('1.0', 'UTF-8');
$writer->startElement('catalog');
foreach ([['id' => 1, 'title' => 'Book A'], ['id' => 2, 'title' => 'Book B']] as $book) {
    $writer->startElement('book');
    $writer->writeAttribute('id', $book['id']);
    $writer->writeElement('title', $book['title']);
    $writer->endElement();
}
$writer->endElement();
$writer->endDocument();
$xml = $writer->outputMemory();
$doc = new DOMDocument();
$doc->loadXML($xml);
echo $doc->getElementsByTagName('book')->length;
"#,
    );
}

// ── XMLReader ────────────────────────────────────────────────

#[test]
fn xml_reader_basic() {
    compile_ok(
        r#"<?php
$xml = '<?xml version="1.0"?><root><item>one</item><item>two</item></root>';
$reader = new XMLReader();
$reader->XML($xml);
$items = [];
while ($reader->read()) {
    if ($reader->nodeType === XMLReader::ELEMENT && $reader->localName === 'item') {
        $reader->read(); // text node
        $items[] = $reader->value;
    }
}
$reader->close();
echo implode(',', $items);
"#,
    );
}

#[test]
fn xml_reader_attributes() {
    compile_ok(
        r#"<?php
$xml = '<items><item id="1" name="A"/><item id="2" name="B"/></items>';
$reader = new XMLReader();
$reader->XML($xml);
$ids = [];
while ($reader->read()) {
    if ($reader->nodeType === XMLReader::ELEMENT && $reader->localName === 'item') {
        $ids[] = $reader->getAttribute('id');
    }
}
$reader->close();
echo implode(',', $ids);
"#,
    );
}

#[test]
fn xml_reader_depth() {
    compile_ok(
        r#"<?php
$xml = '<a><b><c>deep</c></b></a>';
$reader = new XMLReader();
$reader->XML($xml);
$maxDepth = 0;
while ($reader->read()) {
    if ($reader->depth > $maxDepth) $maxDepth = $reader->depth;
}
$reader->close();
echo $maxDepth;
"#,
    );
}

// ── SimpleXML ────────────────────────────────────────────────

#[test]
fn simplexml_basic() {
    compile_ok(
        r#"<?php
$xml = simplexml_load_string('<root><name>Alice</name><age>30</age></root>');
echo $xml->name . ':' . $xml->age;
"#,
    );
}

#[test]
fn simplexml_attributes() {
    compile_ok(
        r#"<?php
$xml = simplexml_load_string('<user id="42" role="admin"><name>Bob</name></user>');
echo $xml['id'] . ':' . $xml['role'] . ':' . $xml->name;
"#,
    );
}

#[test]
fn simplexml_children() {
    compile_ok(
        r#"<?php
$xml = simplexml_load_string('<list><item>a</item><item>b</item><item>c</item></list>');
$count = 0;
foreach ($xml->item as $item) { $count++; }
echo $count;
echo ':' . $xml->item[1];
"#,
    );
}

#[test]
fn simplexml_xpath() {
    compile_ok(
        r#"<?php
$xml = simplexml_load_string('<books><book lang="en"><title>A</title></book><book lang="fr"><title>B</title></book></books>');
$en = $xml->xpath('//book[@lang="en"]/title');
echo count($en) . ':' . $en[0];
"#,
    );
}

#[test]
fn simplexml_to_array() {
    compile_ok(
        r#"<?php
$xml = simplexml_load_string('<data><key>value</key><num>42</num></data>');
$arr = json_decode(json_encode($xml), true);
echo $arr['key'] . ':' . $arr['num'];
"#,
    );
}
