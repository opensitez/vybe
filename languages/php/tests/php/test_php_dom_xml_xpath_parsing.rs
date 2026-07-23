use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: DOM, XML & XPath Parsing — DOMDocument, DOMXPath, DOMElement, DOMNodeList, loadXML, loadHTML, evaluate, query
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_domdocument_load_xml_and_xpath_query() {
    let out = run_prints(
        r#"<?php
$xml = <<<XML
<?xml version="1.0"?>
<catalog>
    <book id="bk101"><title>PHP 8 in Action</title><price>39.95</price></book>
    <book id="bk102"><title>Rust Systems</title><price>49.95</price></book>
</catalog>
XML;

$doc = new DOMDocument();
$doc->loadXML($xml);

$xpath = new DOMXPath($doc);
$titles = $xpath->query("//book/title");

$out = [];
foreach ($titles as $node) {
    $out[] = $node->nodeValue;
}
echo implode(" | ", $out);
"#,
    );
    assert_eq!(out, vec!["PHP 8 in Action | Rust Systems"]);
}

#[test]
fn test_php_domdocument_create_element_and_append() {
    let out = run_prints(
        r#"<?php
$doc = new DOMDocument("1.0", "UTF-8");
$root = $doc->createElement("response");
$status = $doc->createElement("status", "success");
$root->appendChild($status);
$doc->appendChild($root);

echo $doc->saveXML();
"#,
    );
    assert_eq!(
        out,
        vec![
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<response><status>success</status></response>"
        ]
    );
}

#[test]
fn test_php_domdocument_load_html_suppress_errors() {
    compile_ok(
        r#"<?php
$html = '<div class="content"><p>Unclosed paragraph</div>';
$doc = new DOMDocument();
libxml_use_internal_errors(true);
$doc->loadHTML($html);
libxml_clear_errors();

$p = $doc->getElementsByTagName("p")->item(0);
echo $p ? $p->textContent : "NO_P";
"#,
    );
}

#[test]
fn test_php_dom_element_attribute_get_set() {
    compile_ok(
        r#"<?php
$doc = new DOMDocument();
$elem = $doc->createElement("a", "Click Here");
$elem->setAttribute("href", "https://example.com");
$elem->setAttribute("target", "_blank");

echo $elem->getAttribute("href") . " target=" . $elem->getAttribute("target");
"#,
    );
}

#[test]
fn test_php_domxpath_evaluate_expressions() {
    compile_ok(
        r#"<?php
$xml = '<items><item price="10"/><item price="20"/></items>';
$doc = new DOMDocument();
$doc->loadXML($xml);

$xpath = new DOMXPath($doc);
$total = $xpath->evaluate("sum(//item/@price)");
echo "Total: $total";
"#,
    );
}

#[test]
fn test_php_domdocument_import_node_cross_document() {
    compile_ok(
        r#"<?php
$doc1 = new DOMDocument();
$doc1->loadXML("<root><source>Data</source></root>");

$doc2 = new DOMDocument();
$doc2->loadXML("<target/>");

$imported = $doc2->importNode($doc1->documentElement->firstChild, true);
$doc2->documentElement->appendChild($imported);

echo $doc2->saveXML();
"#,
    );
}

#[test]
fn test_php_simplexml_element_parsing() {
    compile_ok(
        r#"<?php
$xmlStr = "<user><name>Alice</name><email>alice@domain.com</email></user>";
$sxml = simplexml_load_string($xmlStr);
echo "{$sxml->name} <{$sxml->email}>";
"#,
    );
}

#[test]
fn test_php_simplexml_to_domdocument_conversion() {
    compile_ok(
        r#"<?php
$sxml = simplexml_load_string("<root><item id='1'/></root>");
$domElem = dom_import_simplexml($sxml);
echo $domElem->nodeName;
"#,
    );
}

#[test]
fn test_php_domnode_remove_child() {
    compile_ok(
        r#"<?php
$doc = new DOMDocument();
$doc->loadXML("<root><child1/><child2/></root>");
$root = $doc->documentElement;
$root->removeChild($root->firstChild);
echo $doc->saveXML();
"#,
    );
}

#[test]
fn test_php_dom_character_data_cdata_section() {
    compile_ok(
        r#"<?php
$doc = new DOMDocument();
$cdata = $doc->createCDATASection("<code>if (a < b)</code>");
$root = $doc->createElement("script");
$root->appendChild($cdata);
$doc->appendChild($root);
echo $doc->saveXML();
"#,
    );
}
