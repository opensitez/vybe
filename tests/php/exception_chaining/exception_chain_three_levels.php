<?php
// vybe-test: php/exception_chaining/exception_chain_three_levels
// origin: languages/php/tests/php/test_exception_chaining.rs

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

try {
    try {
        try { throw new Exception('root'); }
        catch (Exception $e) { throw new RuntimeException('mid', 0, $e); }
    } catch (RuntimeException $e) { throw new LogicException('top', 0, $e); }
} catch (LogicException $e) {
    $chain = [];
    $cur = $e;
    while ($cur !== null) { $chain[] = $cur->getMessage(); $cur = $cur->getPrevious(); }
    echo implode('>', $chain);
}

__vybe_check(ob_get_clean(), "top>mid>root");
