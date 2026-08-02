<?php
// vybe-test: php/php_spl_caching_iterator_lookahead/test_php_spl_caching_iterator_to_string_conversion
// origin: languages/php/tests/php/test_php_spl_caching_iterator_lookahead.rs

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

$arr = new ArrayIterator(["apple", "banana"]);
$it = new CachingIterator($arr, CachingIterator::TOSTRING_USE_CURRENT);

$it->rewind();
echo (string)$it;

__vybe_check(ob_get_clean(), "apple");
