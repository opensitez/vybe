<?php
// vybe-test: php/modern_php_deep/match_multiple_conditions
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

function httpStatus(int $code): string {
    return match($code) {
        200, 201 => "success",
        301, 302 => "redirect",
        404 => "not found",
        500, 502, 503 => "server error",
        default => "unknown" };
}
echo httpStatus(200);
echo httpStatus(301);
echo httpStatus(404);
echo httpStatus(503);
echo httpStatus(418);

__vybe_check(ob_get_clean(), "successredirectnot foundserver errorunknown");
