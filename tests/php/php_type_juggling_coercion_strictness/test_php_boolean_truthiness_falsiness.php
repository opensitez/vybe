<?php
// vybe-test: php/php_type_juggling_coercion_strictness/test_php_boolean_truthiness_falsiness
// origin: languages/php/tests/php/test_php_type_juggling_coercion_strictness.rs

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

$falsyValues = [0, 0.0, "", "0", [], null, false];
$falsyCount = 0;
foreach ($falsyValues as $v) {
    if (!$v) $falsyCount++;
}
echo "Falsy Count: $falsyCount / " . count($falsyValues);

__vybe_check(ob_get_clean(), "Falsy Count: 7 / 7");
