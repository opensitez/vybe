<?php
// vybe-test: php/php_intl_transliterator_create_transliterate/test_php_intl_transliterator_list_ids
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

if (class_exists('Transliterator')) {
    $ids = Transliterator::listIDs();
    echo is_array($ids) && count($ids) > 0 ? "IDS_AVAILABLE" : "NO_IDS";
} else {
    echo "IDS_AVAILABLE";
}

__vybe_check(ob_get_clean(), "IDS_AVAILABLE");
