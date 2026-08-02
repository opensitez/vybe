<?php
// vybe-test: php/intl_formatting/resourcebundle_create_when_available
// origin: languages/php/tests/php/test_intl_formatting.rs

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

if (!class_exists('ResourceBundle')) { echo 'skip'; } else {
    $rb = ResourceBundle::create('en', null, false);
    echo $rb !== null ? 'bundle' : 'null';
}

__vybe_check(ob_get_clean(), "bundle");
