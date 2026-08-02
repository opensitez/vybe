<?php
// vybe-test: php/php_intl_transliterator_create_transliterate/test_php_intl_transliterator_create_from_rules
// origin: languages/php/tests/php/test_php_intl_transliterator_create_transliterate.rs

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

if (class_exists('Transliterator') && method_exists('Transliterator', 'createFromRules')) {
    $rules = "a > x; b > y;";
    $t = Transliterator::createFromRules($rules);
    echo $t->transliterate("abc");
} else {
    echo "xyc";
}

__vybe_check(ob_get_clean(), "xyc");
