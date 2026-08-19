<?php
// vybe-test: php/array_column_advanced/array_column_with_missing_key_is_omitted
// origin: languages/php/tests/php/test_array_column_advanced.rs

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

$rows = [
    ['id' => 1, 'name' => 'Alice'],
    ['id' => 2],
    ['id' => 3, 'name' => null],
];
$names = array_column($rows, 'name');
echo implode('|', $names);

__vybe_check(ob_get_clean(), "Alice|");
