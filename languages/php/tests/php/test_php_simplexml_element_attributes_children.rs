use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP SimpleXML: SimpleXMLElement Property Access, Attributes & Children
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_simplexml_load_string_property_access() {
    let out = run_prints(
        r##"<?php
$xml = "<user><name>John</name><email>john@example.com</email></user>";
$sxe = simplexml_load_string($xml);
echo "Name=" . $sxe->name . " Email=" . $sxe->email;
"##,
    );
    assert_eq!(out, vec!["Name=John Email=john@example.com"]);
}

#[test]
fn test_php_simplexml_attribute_array_access() {
    let out = run_prints(
        r##"<?php
$xml = '<item id="42" category="books">PHP Handbook</item>';
$sxe = simplexml_load_string($xml);
echo "ID=" . $sxe["id"] . " Cat=" . $sxe["category"];
"##,
    );
    assert_eq!(out, vec!["ID=42 Cat=books"]);
}

#[test]
fn test_php_simplexml_children_iteration() {
    let out = run_prints(
        r##"<?php
$xml = '<menu><food>Pizza</food><food>Burger</food><food>Tacos</food></menu>';
$sxe = simplexml_load_string($xml);

$foods = [];
foreach ($sxe->children() as $food) {
    $foods[] = (string)$food;
}
echo implode(",", $foods);
"##,
    );
    assert_eq!(out, vec!["Pizza,Burger,Tacos"]);
}

#[test]
fn test_php_simplexml_add_child_element() {
    compile_ok(
        r##"<?php
$sxe = new SimpleXMLElement("<root/>");
$child = $sxe->addChild("setting", "enabled");
echo $sxe->setting == "enabled" ? "ADD_CHILD_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_simplexml_add_attribute() {
    compile_ok(
        r##"<?php
$sxe = new SimpleXMLElement("<product/>");
$sxe->addAttribute("price", "19.99");
echo (string)$sxe["price"] === "19.99" ? "ADD_ATTR_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_simplexml_as_xml_export() {
    compile_ok(
        r##"<?php
$sxe = new SimpleXMLElement("<note><to>Tove</to></note>");
$xmlOut = $sxe->asXML();
echo str_contains($xmlOut, "<to>Tove</to>") ? "AS_XML_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_simplexml_xpath_search() {
    compile_ok(
        r##"<?php
$xml = '<store><book price="10"/><book price="20"/></store>';
$sxe = simplexml_load_string($xml);
$res = $sxe->xpath("//book[@price='20']");
echo count($res) === 1 ? "SIMPLEXML_XPATH_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_simplexml_attributes_iterator() {
    compile_ok(
        r##"<?php
$xml = '<tag a="1" b="2"/>';
$sxe = simplexml_load_string($xml);
$attrs = [];
foreach ($sxe->attributes() as $name => $val) {
    $attrs[$name] = (string)$val;
}
echo isset($attrs["a"]) && isset($attrs["b"]) ? "ATTRS_ITERATOR_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_simplexml_dom_import_simplexml() {
    compile_ok(
        r##"<?php
$sxe = new SimpleXMLElement("<data><val>123</val></data>");
$dom = dom_import_simplexml($sxe);
echo $dom instanceof DOMElement ? "DOM_IMPORT_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_simplexml_count_elements() {
    compile_ok(
        r##"<?php
$xml = '<items><i/><i/><i/></items>';
$sxe = simplexml_load_string($xml);
echo $sxe->count() === 3 ? "SIMPLEXML_COUNT_3_OK" : "FAIL";
"##,
    );
}
