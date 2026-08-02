<?php
// vybe-test: php/spread_operator/spread_array_merge_with_numeric_reindexing
// origin: languages/php/tests/php/test_spread_operator.rs

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

$a = [10 => 'a'];
$b = [11 => 'b'];
$c = [...$a, ...$b];
echo count($c);
echo '|';
echo isset($c[10]) ? 'has10' : 'no10';
echo '|';
echo array_key_exists(0, $c) ? 'has0' : 'no0';
echo '|';
echo array_key_exists(1, $c) ? 'has1' : 'no1';

__vybe_check(ob_get_clean(), "2|no10|has0|has1");
