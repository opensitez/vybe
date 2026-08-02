<?php
// vybe-test: php/generator_errors/generator_nested_yield_from_with_return_in_middle
// origin: languages/php/tests/php/test_generator_errors.rs

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

function mid(): Generator { yield 'm'; return 'R'; }
function top(): Generator { yield 't'; $r = yield from mid(); yield $r; }
echo implode(',', iterator_to_array(top()));

__vybe_check(ob_get_clean(), "m,R");
