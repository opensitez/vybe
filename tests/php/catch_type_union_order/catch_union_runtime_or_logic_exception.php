<?php
// vybe-test: php/catch_type_union_order/catch_union_runtime_or_logic_exception
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

function boom(int $n): void {
    if ($n === 1) { throw new RuntimeException('run'); }
    throw new LogicException('logic');
}
foreach ([1, 2] as $n) {
    try { boom($n); }
    catch (RuntimeException | LogicException $e) { echo $e->getMessage(); }
}

__vybe_check(ob_get_clean(), "runlogic");
