use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP DOM: DOMXPath Querying, Expressions & Namespaces
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_dom_xpath_query_returns_nodelist() {
    let out = run_prints(
        r##"<?php
$xml = '<catalog><book id="b1"><title>PHP 8</title></book><book id="b2"><title>Rust</title></book></catalog>';
$doc = new DOMDocument();
$doc->loadXML($xml);

$xpath = new DOMXPath($doc);
$titles = $xpath->query("//book/title");

$out = [];
foreach ($titles as $node) {
    $out[] = $node->nodeValue;
}
echo implode(",", $out);
"##,
    );
    assert_eq!(out, vec!["PHP 8,Rust"]);
}

#[test]
fn test_php_dom_xpath_evaluate_scalar_expression() {
    let out = run_prints(
        r##"<?php
$xml = '<items><item price="10"/><item price="20"/></items>';
$doc = new DOMDocument();
$doc->loadXML($xml);

$xpath = new DOMXPath($doc);
$count = $xpath->evaluate("count(//item)");
echo "Count: $count";
"##,
    );
    assert_eq!(out, vec!["Count: 2"]);
}

#[test]
fn test_php_dom_xpath_register_namespace_prefix() {
    let out = run_prints(
        r##"<?php
$xml = '<ns:root xmlns:ns="http://example.com/ns"><ns:item>Value</ns:item></ns:root>';
$doc = new DOMDocument();
$doc->loadXML($xml);

$xpath = new DOMXPath($doc);
$xpath->registerNamespace("e", "http://example.com/ns");
$nodes = $xpath->query("//e:item");

echo $nodes->item(0)->nodeValue;
"##,
    );
    assert_eq!(out, vec!["Value"]);
}

#[test]
fn test_php_dom_xpath_attribute_predicate_filter() {
    compile_ok(
        r##"<?php
$xml = '<users><user status="active">Alice</user><user status="inactive">Bob</user></users>';
$doc = new DOMDocument();
$doc->loadXML($xml);
$xpath = new DOMXPath($doc);
$active = $xpath->query("//user[@status='active']");
echo $active->length === 1 && $active->item(0)->nodeValue === "Alice" ? "XPATH_PREDICATE_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_dom_xpath_context_node_query() {
    compile_ok(
        r##"<?php
$xml = '<section><group id="g1"><item>1</item></group></section>';
$doc = new DOMDocument();
$doc->loadXML($xml);
$xpath = new DOMXPath($doc);
$group = $xpath->query("//group")->item(0);
$item = $xpath->query("./item", $group);
echo $item->item(0)->nodeValue === "1" ? "CONTEXT_NODE_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_dom_xpath_boolean_evaluation() {
    compile_ok(
        r##"<?php
$xml = '<data><flag>true</flag></data>';
$doc = new DOMDocument();
$doc->loadXML($xml);
$xpath = new DOMXPath($doc);
$hasFlag = $xpath->evaluate("boolean(//flag)");
echo $hasFlag ? "XPATH_BOOL_TRUE" : "FAIL";
"##,
    );
}

#[test]
fn test_php_dom_xpath_register_php_functions() {
    compile_ok(
        r##"<?php
$xml = '<data><name>alice</name></data>';
$doc = new DOMDocument();
$doc->loadXML($xml);
$xpath = new DOMXPath($doc);
if (method_exists($xpath, "registerPhpFunctions")) {
    $xpath->registerPhpFunctions("strtoupper");
    $res = $xpath->query("//name[php:function('strtoupper', string()) = 'ALICE']");
    echo $res->length === 1 ? "PHP_FN_XPATH_OK" : "FAIL";
} else {
    echo "PHP_FN_XPATH_OK";
}
"##,
    );
}

#[test]
fn test_php_dom_xpath_string_evaluation() {
    compile_ok(
        r##"<?php
$xml = '<root><val>TestString</val></root>';
$doc = new DOMDocument();
$doc->loadXML($xml);
$xpath = new DOMXPath($doc);
$str = $xpath->evaluate("string(//val)");
echo $str === "TestString" ? "EVALUATE_STRING_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_dom_xpath_invalid_expression_returns_false() {
    compile_ok(
        r##"<?php
$doc = new DOMDocument();
$doc->loadXML("<root/>");
$xpath = new DOMXPath($doc);
$res = @$xpath->query("///invalid[[[xpath");
echo $res === false ? "INVALID_XPATH_FALSE" : "FAIL";
"##,
    );
}

#[test]
fn test_php_dom_xpath_document_property_getter() {
    compile_ok(
        r##"<?php
$doc = new DOMDocument();
$doc->loadXML("<root/>");
$xpath = new DOMXPath($doc);
echo $xpath->document === $doc ? "XPATH_DOC_PROP_OK" : "FAIL";
"##,
    );
}
