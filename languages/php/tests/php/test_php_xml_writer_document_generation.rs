use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP XMLWriter: Document Generation, Formatting & Streams
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_xml_writer_memory_generation() {
    let out = run_prints(
        r##"<?php
$w = new XMLWriter();
$w->openMemory();
$w->startDocument("1.0", "UTF-8");
$w->startElement("catalog");
$w->writeElement("item", "Laptop");
$w->endElement();
$w->endDocument();

echo trim($w->outputMemory());
"##,
    );
    assert_eq!(
        out,
        vec!["<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<catalog><item>Laptop</item></catalog>"]
    );
}

#[test]
fn test_php_xml_writer_write_attribute_formatting() {
    let out = run_prints(
        r##"<?php
$w = new XMLWriter();
$w->openMemory();
$w->startElement("user");
$w->writeAttribute("id", "42");
$w->writeAttribute("status", "active");
$w->text("User Details");
$w->endElement();

echo $w->outputMemory();
"##,
    );
    assert_eq!(
        out,
        vec!["<user id=\"42\" status=\"active\">User Details</user>"]
    );
}

#[test]
fn test_php_xml_writer_indentation_formatting() {
    let out = run_prints(
        r##"<?php
$w = new XMLWriter();
$w->openMemory();
$w->setIndent(true);
$w->setIndentString("  ");
$w->startElement("root");
$w->writeElement("child", "val");
$w->endElement();

echo str_contains($w->outputMemory(), "  <child>") ? "INDENTED_OK" : "PLAIN";
"##,
    );
    assert_eq!(out, vec!["INDENTED_OK"]);
}

#[test]
fn test_php_xml_writer_cdata_element() {
    compile_ok(
        r##"<?php
$w = new XMLWriter();
$w->openMemory();
$w->startElement("script");
$w->writeCdata("function foo() { return a < b; }");
$w->endElement();
$xml = $w->outputMemory();
echo str_contains($xml, "<![CDATA[") ? "WRITE_CDATA_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_xml_writer_comment_element() {
    compile_ok(
        r##"<?php
$w = new XMLWriter();
$w->openMemory();
$w->writeComment("This is an XML comment");
$w->startElement("tag");
$w->endElement();
$xml = $w->outputMemory();
echo str_contains($xml, "<!--This is an XML comment-->") ? "WRITE_COMMENT_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_xml_writer_dtd_element() {
    compile_ok(
        r##"<?php
$w = new XMLWriter();
$w->openMemory();
$w->startDtd("html");
$w->endDtd();
$xml = $w->outputMemory();
echo str_contains($xml, "<!DOCTYPE html") ? "WRITE_DTD_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_xml_writer_start_end_attribute() {
    compile_ok(
        r##"<?php
$w = new XMLWriter();
$w->openMemory();
$w->startElement("link");
$w->startAttribute("href");
$w->text("https://example.com");
$w->endAttribute();
$w->endElement();
$xml = $w->outputMemory();
echo str_contains($xml, 'href="https://example.com"') ? "START_END_ATTR_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_xml_writer_namespace_element() {
    compile_ok(
        r##"<?php
$w = new XMLWriter();
$w->openMemory();
$w->startElementNs("ns", "element", "http://example.com/ns");
$w->text("Namespace Content");
$w->endElement();
$xml = $w->outputMemory();
echo str_contains($xml, "ns:element") && str_contains($xml, "http://example.com/ns") ? "ELEMENT_NS_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_xml_writer_flush_output_memory() {
    compile_ok(
        r##"<?php
$w = new XMLWriter();
$w->openMemory();
$w->writeElement("a", "b");
$chunk1 = $w->flush();
$w->writeElement("c", "d");
$chunk2 = $w->flush();
echo str_contains($chunk1, "<a>b</a>") && str_contains($chunk2, "<c>d</c>") ? "FLUSH_MEMORY_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_xml_writer_pi_processing_instruction() {
    compile_ok(
        r##"<?php
$w = new XMLWriter();
$w->openMemory();
$w->writePi("php-stylesheet", 'href="style.css"');
$w->startElement("page");
$w->endElement();
$xml = $w->outputMemory();
echo str_contains($xml, "<?php-stylesheet") ? "WRITE_PI_OK" : "FAIL";
"##,
    );
}
