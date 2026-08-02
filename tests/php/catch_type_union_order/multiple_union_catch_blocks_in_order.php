<?php
// vybe-test: php/catch_type_union_order/multiple_union_catch_blocks_in_order
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

function step(int $n): void {
    if ($n === 1) { throw new TypeError('t'); }
    if ($n === 2) { throw new ValueError('v'); }
    throw new RuntimeException('r');
}
foreach ([1, 2, 3] as $n) {
    try { step($n); }
    catch (TypeError | ValueError $e) { echo 'tv'; }
    catch (RuntimeException $e) { echo 'rt'; }
}

__vybe_check(ob_get_clean(), "tvtvrt");
