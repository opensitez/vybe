use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP XMLReader: Streaming XML Parser & Node Inspection
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_xml_reader_read_nodes_sequentially() {
    let out = run_prints(
        r##"<?php
$xml = '<users><user>Alice</user><user>Bob</user></users>';
$reader = new XMLReader();
$reader->xml($xml);

$names = [];
while ($reader->read()) {
    if ($reader->nodeType === XMLReader::ELEMENT && $reader->name === "user") {
        $reader->read(); // Read inner text node
        $names[] = $reader->value;
    }
}
$reader->close();
echo implode(",", $names);
"##,
    );
    assert_eq!(out, vec!["Alice,Bob"]);
}

#[test]
fn test_php_xml_reader_get_attribute_by_name() {
    let out = run_prints(
        r##"<?php
$xml = '<item id="999" status="active"/>';
$reader = new XMLReader();
$reader->xml($xml);

while ($reader->read()) {
    if ($reader->nodeType === XMLReader::ELEMENT) {
        echo "ID=" . $reader->getAttribute("id") . " Status=" . $reader->getAttribute("status");
    }
}
$reader->close();
"##,
    );
    assert_eq!(out, vec!["ID=999 Status=active"]);
}

#[test]
fn test_php_xml_reader_read_outer_xml_string() {
    let out = run_prints(
        r##"<?php
$xml = '<root><sub id="1">Text</sub></root>';
$reader = new XMLReader();
$reader->xml($xml);

while ($reader->read()) {
    if ($reader->nodeType === XMLReader::ELEMENT && $reader->name === "sub") {
        echo $reader->readOuterXML();
    }
}
$reader->close();
"##,
    );
    assert_eq!(out, vec!["<sub id=\"1\">Text</sub>"]);
}

#[test]
fn test_php_xml_reader_expand_to_dom_node() {
    compile_ok(
        r##"<?php
$xml = '<node attr="val">Content</node>';
$reader = new XMLReader();
$reader->xml($xml);
$reader->read();
$domNode = $reader->expand();
echo $domNode instanceof DOMNode ? "EXPAND_TO_DOM_OK" : "FAIL";
$reader->close();
"##,
    );
}

#[test]
fn test_php_xml_reader_next_skips_children() {
    compile_ok(
        r##"<?php
$xml = '<list><group><item/></group><target/></list>';
$reader = new XMLReader();
$reader->xml($xml);
$reader->read(); // list
$reader->read(); // group
$reader->next("target"); // Skip group children to target
echo $reader->name === "target" ? "NEXT_TARGET_OK" : "FAIL";
$reader->close();
"##,
    );
}

#[test]
fn test_php_xml_reader_is_empty_element_property() {
    compile_ok(
        r##"<?php
$xml = '<container><empty/><nonempty>text</nonempty></container>';
$reader = new XMLReader();
$reader->xml($xml);
$reader->read(); // container
$reader->read(); // empty
echo $reader->isEmptyElement ? "IS_EMPTY_ELEMENT_TRUE" : "FAIL";
$reader->close();
"##,
    );
}

#[test]
fn test_php_xml_reader_depth_property() {
    compile_ok(
        r##"<?php
$xml = '<level0><level1><level2/></level1></level0>';
$reader = new XMLReader();
$reader->xml($xml);
$reader->read(); // level0
$reader->read(); // level1
echo $reader->depth === 1 ? "DEPTH_1_OK" : "FAIL";
$reader->close();
"##,
    );
}

#[test]
fn test_php_xml_reader_set_parser_property_option() {
    compile_ok(
        r##"<?php
$reader = new XMLReader();
$reader->setParserProperty(XMLReader::SUBST_ENTITIES, true);
echo $reader->getParserProperty(XMLReader::SUBST_ENTITIES) ? "PARSER_PROP_OK" : "FAIL";
$reader->close();
"##,
    );
}

#[test]
fn test_php_xml_reader_read_inner_xml_string() {
    compile_ok(
        r##"<?php
$xml = '<wrapper><content>Hello XMLReader</content></wrapper>';
$reader = new XMLReader();
$reader->xml($xml);
$reader->read(); // wrapper
echo str_contains($reader->readInnerXML(), "<content>") ? "INNER_XML_OK" : "FAIL";
$reader->close();
"##,
    );
}

#[test]
fn test_php_xml_reader_move_to_attribute_no() {
    compile_ok(
        r##"<?php
$xml = '<element a="valA" b="valB"/>';
$reader = new XMLReader();
$reader->xml($xml);
$reader->read();
$reader->moveToAttributeNo(1);
echo $reader->name === "b" && $reader->value === "valB" ? "MOVE_TO_ATTR1_OK" : "FAIL";
$reader->close();
"##,
    );
}
