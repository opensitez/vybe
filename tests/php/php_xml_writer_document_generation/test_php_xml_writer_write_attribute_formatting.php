<?php
// vybe-test: php/php_xml_writer_document_generation/test_php_xml_writer_write_attribute_formatting
// origin: languages/php/tests/php/test_php_xml_writer_document_generation.rs

function __vybe_check($got, $want) {
    // Match the Rust harness's normalisation: strip \r, then drop trailing
    // newlines (it split on "\n" and popped empty trailing elements).
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    // Replay the program's own output so running the file by hand still
    // behaves like the program it was extracted from.
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

$w = new XMLWriter();
$w->openMemory();
$w->startElement("user");
$w->writeAttribute("id", "42");
$w->writeAttribute("status", "active");
$w->text("User Details");
$w->endElement();

echo $w->outputMemory();

__vybe_check(ob_get_clean(), "<user id=\"42\" status=\"active\">User Details</user>");
