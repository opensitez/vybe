<?php
// vybe-test: php/php_constants_extended/sort_flags_defined
// origin: languages/php/tests/php/test_php_constants_extended.rs

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

echo defined('SORT_REGULAR') ? '1' : '0';
echo defined('SORT_NUMERIC') ? '1' : '0';
echo defined('SORT_STRING') ? '1' : '0';
echo defined('SORT_NATURAL') ? '1' : '0';

__vybe_check(ob_get_clean(), "1111");
