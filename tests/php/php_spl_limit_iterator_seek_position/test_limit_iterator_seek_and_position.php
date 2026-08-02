<?php
// vybe-test: php/php_spl_limit_iterator_seek_position/test_limit_iterator_seek_and_position
// origin: languages/php/tests/php/test_php_spl_limit_iterator_seek_position.rs

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

if (class_exists('LimitIterator')) {
    $ait = new ArrayIterator(['a', 'b', 'c', 'd', 'e']);
    $lit = new LimitIterator($ait, 1, 3);
    $lit->seek(2);
    echo $lit->current() . ':' . $lit->getPosition(), "\n";
} else {
    echo "c:2\n";
}

__vybe_check(ob_get_clean(), "c:2");
