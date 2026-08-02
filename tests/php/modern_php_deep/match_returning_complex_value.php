<?php
// vybe-test: php/modern_php_deep/match_returning_complex_value
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

function getConfig(string $env): array {
    return match($env) {
        "dev"  => ["debug" => true,  "log" => "verbose"],
        "prod" => ["debug" => false, "log" => "error"],
        default => ["debug" => false, "log" => "warning"],
    };
}
$cfg = getConfig("dev");
echo $cfg["log"];
$cfg2 = getConfig("prod");
echo $cfg2["debug"] ? "debug" : "no-debug";

__vybe_check(ob_get_clean(), "verboseno-debug");
