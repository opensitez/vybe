<?php
// vybe-test: php/php_expressions_match_nullsafe/test_php_match_expression_strict_type_comparison
// origin: languages/php/tests/php/test_php_expressions_match_nullsafe.rs

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

$input = "10";
$res = match ($input) {
    10 => "integer",
    "10" => "string",
    default => "other",
};
echo $res;

__vybe_check(ob_get_clean(), "string");
