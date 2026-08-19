<?php
// vybe-test: php/php_xml_writer_document_generation/test_php_xml_writer_flush_output_memory
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
$w->writeElement("a", "b");
$chunk1 = $w->flush();
$w->writeElement("c", "d");
$chunk2 = $w->flush();
echo str_contains($chunk1, "<a>b</a>") && str_contains($chunk2, "<c>d</c>") ? "FLUSH_MEMORY_OK" : "FAIL";


__vybe_check(ob_get_clean(), "FLUSH_MEMORY_OK");
