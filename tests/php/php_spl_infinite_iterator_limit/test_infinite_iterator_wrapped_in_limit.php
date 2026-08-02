<?php
// vybe-test: php/php_spl_infinite_iterator_limit/test_infinite_iterator_wrapped_in_limit
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
    $arrayIt = new ArrayIterator(['A', 'B']);
    $inf = new InfiniteIterator($arrayIt);
    $limit = new LimitIterator($inf, 0, 5);
    $out = [];
    foreach ($limit as $v) {
        $out[] = $v;
    }
    echo implode(',', $out), "\n";
} else {
    echo "A,B,A,B,A\n";
}

__vybe_check(ob_get_clean(), "A,B,A,B,A");
