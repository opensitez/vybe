<?php
// vybe-test: php/php_xml_reader_streaming_parser/test_php_xml_reader_read_nodes_sequentially
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

__vybe_check(ob_get_clean(), "Alice,Bob");
