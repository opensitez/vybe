<?php
// vybe-test: php/php_array_key_exists_search_in_array/test_php_array_values_reindexes_keys
// origin: languages/php/tests/php/test_php_array_key_exists_search_in_array.rs

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

$arr = [10 => "x", 20 => "y", 30 => "z"];
$reindexed = array_values($arr);
echo implode(":", array_keys($reindexed)) . " = " . implode(",", $reindexed);

__vybe_check(ob_get_clean(), "0:1:2 = x,y,z");
