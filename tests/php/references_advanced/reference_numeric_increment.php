<?php
// vybe-test: php/references_advanced/reference_numeric_increment
// origin: languages/php/tests/php/test_references_advanced.rs

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

$counters = ['a' => 0, 'b' => 0];
$ref = &$counters['a'];
for ($i = 0; $i < 5; $i++) $ref++;
echo $counters['a'] . ',' . $counters['b'];

__vybe_check(ob_get_clean(), "5,0");
