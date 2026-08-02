<?php
// vybe-test: php/catch_type_union_order/catch_exception_chain_with_union_and_fallback
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

function work(int $step): void {
    if ($step === 1) { throw new InvalidArgumentException('a'); }
    if ($step === 2) { throw new UnexpectedValueException('u'); }
    throw new RuntimeException('r');
}
foreach ([1, 2, 3] as $s) {
    try { work($s); }
    catch (InvalidArgumentException | UnexpectedValueException $e) { echo 'iu'; }
    catch (RuntimeException $e) { echo 'rt'; }
}

__vybe_check(ob_get_clean(), "iuiurt");
