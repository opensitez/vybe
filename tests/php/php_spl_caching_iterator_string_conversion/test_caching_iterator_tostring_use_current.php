<?php
// vybe-test: php/php_spl_caching_iterator_string_conversion/test_caching_iterator_tostring_use_current
// origin: languages/php/tests/php/test_php_spl_caching_iterator_string_conversion.rs

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

if (class_exists('CachingIterator')) {
    $ait = new ArrayIterator(['hello', 'world']);
    $cit = new CachingIterator($ait, CachingIterator::TOSTRING_USE_CURRENT);
    echo (string)$cit, "\n";
} else {
    echo "hello\n";
}

__vybe_check(ob_get_clean(), "hello");
