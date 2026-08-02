<?php
// vybe-test: php/union_types_runtime/intersection_iterator_and_countable
// origin: languages/php/tests/php/test_union_types_runtime.rs

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

function len(Iterator&Countable $c): int { return count($c); }
$it = new ArrayIterator([1, 2, 3]);
echo len($it);

__vybe_check(ob_get_clean(), "3");
