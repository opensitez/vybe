<?php
// vybe-test: php/php_xml_reader_streaming_parser/test_php_xml_reader_read_outer_xml_string
// origin: languages/php/tests/php/test_php_xml_reader_streaming_parser.rs

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

$xml = '<root><sub id="1">Text</sub></root>';
$reader = new XMLReader();
$reader->xml($xml);

while ($reader->read()) {
    if ($reader->nodeType === XMLReader::ELEMENT && $reader->name === "sub") {
        echo $reader->readOuterXML();
    }
}
$reader->close();

__vybe_check(ob_get_clean(), "<sub id=\"1\">Text</sub>");
