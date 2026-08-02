<?php
// vybe-test: php/arrays/array_column_with_no_match_returns_null
// origin: languages/php/tests/php/test_arrays.rs

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

$rows = [['id' => 1, 'name' => 'A'], ['id' => 2]];
$vals = array_column($rows, 'name');
$lookup = array_column($rows, 'name', 'id');
echo count($vals) . '|' . (array_key_exists(3, $lookup) ? 'has' : 'no');

__vybe_check(ob_get_clean(), "1|no");
