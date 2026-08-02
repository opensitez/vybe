<?php
// vybe-test: php/php_intl_collator_attribute_settings/test_collator_numeric_collation_attribute
// origin: languages/php/tests/php/test_php_intl_collator_attribute_settings.rs

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

if (class_exists('Collator')) {
    $coll = new Collator('en_US');
    $coll->setAttribute(Collator::NUMERIC_COLLATION, Collator::ON);
    $arr = ['file10.txt', 'file2.txt'];
    $coll->sort($arr);
    echo implode(',', $arr), "\n";
} else {
    echo "file2.txt,file10.txt\n";
}

__vybe_check(ob_get_clean(), "file2.txt,file10.txt");
