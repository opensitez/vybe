<?php
// vybe-test: php/catch_type_union_order/catch_parse_error_or_type_error_union
// origin: languages/php/tests/php/test_catch_type_union_order.rs

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

function fail(int $mode): void {
    if ($mode === 1) { throw new ParseError('parse'); }
    throw new TypeError('type');
}
foreach ([1, 2] as $m) {
    try { fail($m); }
    catch (ParseError | TypeError $e) { echo $e->getMessage(); }
}

__vybe_check(ob_get_clean(), "parsetype");
