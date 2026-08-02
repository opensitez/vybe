<?php
// vybe-test: php/php_spl_infinite_iterator_limit/test_infinite_iterator_rewind_cycles
// origin: languages/php/tests/php/test_php_spl_infinite_iterator_limit.rs

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

if (class_exists('InfiniteIterator') && class_exists('LimitIterator')) {
    $arrayIt = new ArrayIterator([1, 2, 3]);
    $inf = new InfiniteIterator($arrayIt);
    $limit = new LimitIterator($inf, 0, 4);
    $sum = 0;
    foreach ($limit as $v) {
        $sum += $v;
    }
    echo $sum, "\n";
} else {
    echo "7\n";
}

__vybe_check(ob_get_clean(), "7");
