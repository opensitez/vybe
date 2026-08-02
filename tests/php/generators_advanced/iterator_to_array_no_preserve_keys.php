<?php
// vybe-test: php/generators_advanced/iterator_to_array_no_preserve_keys
// origin: languages/php/tests/php/test_generators_advanced.rs

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

function words() {
    yield 5 => "apple";
    yield 3 => "banana";
    yield 7 => "cherry";
}
$arr = iterator_to_array(words(), false);
echo implode(",", $arr);

__vybe_check(ob_get_clean(), "apple,banana,cherry");
