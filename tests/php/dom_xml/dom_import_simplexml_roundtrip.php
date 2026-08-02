<?php
// vybe-test: php/dom_xml/dom_import_simplexml_roundtrip
// origin: languages/php/tests/php/test_dom_xml.rs

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

$sx = simplexml_load_string('<wrap><x>1</x></wrap>');
$doc = new DOMDocument();
$imported = $doc->importNode(dom_import_simplexml($sx), true);
$doc->appendChild($imported);
echo $doc->getElementsByTagName('x')->item(0)->textContent;

__vybe_check(ob_get_clean(), "1");
