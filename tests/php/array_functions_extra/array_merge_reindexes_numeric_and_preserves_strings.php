<?php
// vybe-test: php/array_functions_extra/array_merge_reindexes_numeric_and_preserves_strings
// origin: languages/php/tests/php/test_array_functions_extra.rs

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

$a = [0 => 'a', 3 => 'b', 'k' => 'v'];
$b = ['c', 'd'];
$m = array_merge($a, $b);
echo implode(',', $m);
echo '|';
echo $m['k'];

__vybe_check(ob_get_clean(), "a,b,v,c,d|v");
