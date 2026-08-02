<?php
// vybe-test: php/generator_send_runtime/generator_return_value_after_yield_from
// origin: languages/php/tests/php/test_generator_send_runtime.rs

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

function inner(): Generator { yield 1; return 'done'; }
function outer(): Generator { $r = yield from inner(); return $r; }
$g = outer();
iterator_to_array($g);
echo $g->getReturn();

__vybe_check(ob_get_clean(), "done");
