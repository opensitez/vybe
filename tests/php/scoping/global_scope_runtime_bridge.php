<?php
// vybe-test: php/scoping/global_scope_runtime_bridge
// origin: languages/php/tests/php/test_scoping.rs

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

$count = 10;
function increment_global(): int {
    global $count;
    return ++$count;
}
echo $count . '|';
echo increment_global() . '|';
echo $count;

__vybe_check(ob_get_clean(), "10|11|11");
