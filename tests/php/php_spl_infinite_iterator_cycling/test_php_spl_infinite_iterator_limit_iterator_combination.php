<?php
// vybe-test: php/php_spl_infinite_iterator_cycling/test_php_spl_infinite_iterator_limit_iterator_combination
// origin: languages/php/tests/php/test_php_spl_infinite_iterator_cycling.rs

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

$arr = new ArrayIterator([1, 2, 3]);
$inf = new InfiniteIterator($arr);
$limited = new LimitIterator($inf, 0, 7);

$vals = iterator_to_array($limited, false);
echo implode("-", $vals);

__vybe_check(ob_get_clean(), "1-2-3-1-2-3-1");
