<?php
// vybe-test: php/php_array_chunk_combine_count_values/test_php_array_udiff_with_comparator_value_and_key
// origin: languages/php/tests/php/test_php_array_chunk_combine_count_values.rs

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

$a1 = ["a" => 1, "b" => 2, "c" => 3];
$a2 = ["a" => 1, "B" => 2, "c" => 4];
$diff = array_udiff_assoc(
    $a1,
    $a2,
    fn($v1, $v2) => $v1 <=> $v2
);
echo count($diff) . "|" . implode(",", $diff);

__vybe_check(ob_get_clean(), "2|2,3");
