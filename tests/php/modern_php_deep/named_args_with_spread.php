<?php
// vybe-test: php/modern_php_deep/named_args_with_spread
// origin: languages/php/tests/php/test_modern_php_deep.rs

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

function formatDate(int $year, int $month, int $day): string {
    return sprintf("%04d-%02d-%02d", $year, $month, $day);
}
$params = ["month" => 6, "day" => 15, "year" => 2024];
echo formatDate(...$params);

__vybe_check(ob_get_clean(), "2024-06-15");
