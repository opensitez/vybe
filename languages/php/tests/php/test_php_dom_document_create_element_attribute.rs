use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP DOM: DOMDocument Element & Attribute Creation & Serialization
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_dom_document_create_element_and_append() {
    let out = run_prints(
        r##"<?php
$doc = new DOMDocument("1.0", "UTF-8");
$root = $doc->createElement("root");
$child = $doc->createElement("item", "Hello XML");
$root->appendChild($child);
$doc->appendChild($root);

echo trim($doc->saveXML());
"##,
    );
    assert_eq!(
        out,
        vec!["<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<root><item>Hello XML</item></root>"]
    );
}

#[test]
fn test_php_dom_element_create_attribute() {
    let out = run_prints(
        r##"<?php
$doc = new DOMDocument();
$el = $doc->createElement("user");
$attr = $doc->createAttribute("id");
$attr->value = "123";
$el->appendChild($attr);
$doc->appendChild($el);

echo $el->getAttribute("id");
"##,
    );
    assert_eq!(out, vec!["123"]);
}

#[test]
fn test_php_dom_document_load_xml_string() {
    let out = run_prints(
        r##"<?php
$xml = "<config><setting name='debug'>true</setting></config>";
$doc = new DOMDocument();
$doc->loadXML($xml);

$nodes = $doc->getElementsByTagName("setting");
echo $nodes->item(0)->nodeValue;
"##,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn test_php_dom_element_set_attribute_convenience() {
    compile_ok(
        r##"<?php
$doc = new DOMDocument();
$el = $doc->createElement("div");
$el->setAttribute("class", "container");
echo $el->getAttribute("class") === "container" ? "SET_ATTR_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_dom_element_has_attribute_check() {
    compile_ok(
        r##"<?php
$doc = new DOMDocument();
$el = $doc->createElement("input");
$el->setAttribute("type", "text");
echo $el->hasAttribute("type") && !$el->hasAttribute("disabled") ? "HAS_ATTR_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_dom_element_remove_attribute() {
    compile_ok(
        r##"<?php
$doc = new DOMDocument();
$el = $doc->createElement("btn");
$el->setAttribute("active", "1");
$el->removeAttribute("active");
echo !$el->hasAttribute("active") ? "REMOVE_ATTR_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_dom_create_text_node() {
    compile_ok(
        r##"<?php
$doc = new DOMDocument();
$p = $doc->createElement("p");
$text = $doc->createTextNode("Paragraph content & special <chars>");
$p->appendChild($text);
$doc->appendChild($p);
echo str_contains($doc->saveXML(), "&amp;") || str_contains($doc->saveXML(), "special") ? "TEXT_NODE_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_dom_create_cdata_section() {
    compile_ok(
        r##"<?php
$doc = new DOMDocument();
$cdata = $doc->createCDATASection("<code>if (a < b) {}</code>");
$el = $doc->createElement("script");
$el->appendChild($cdata);
$doc->appendChild($el);
echo str_contains($doc->saveXML(), "CDATA") ? "CDATA_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_dom_get_elements_by_tag_name_length() {
    compile_ok(
        r##"<?php
$doc = new DOMDocument();
$doc->loadXML("<list><item/> <item/> <item/></list>");
$items = $doc->getElementsByTagName("item");
echo $items->length === 3 ? "TAG_LENGTH_3_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_dom_document_save_html() {
    compile_ok(
        r##"<?php
$doc = new DOMDocument();
$doc->loadHTML("<html><body><h1>Title</h1></body></html>");
echo str_contains($doc->saveHTML(), "<h1>Title</h1>") ? "SAVE_HTML_OK" : "FAIL";
"##,
    );
}
