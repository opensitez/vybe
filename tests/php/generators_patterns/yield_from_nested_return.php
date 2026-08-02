<?php
// vybe-test: php/generators_patterns/yield_from_nested_return
// origin: languages/php/tests/php/test_generators_patterns.rs

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

function inner2(): Generator { yield 1; yield 2; return 'inner_done'; }
function outer2(): Generator {
    $result = yield from inner2();
    echo "inner returned: $result\n";
    yield 3;
}
$g = outer2();
iterator_to_array($g);

__vybe_check(ob_get_clean(), "inner returned: inner_done");
