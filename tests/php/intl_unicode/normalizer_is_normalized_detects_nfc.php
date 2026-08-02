<?php
// vybe-test: php/intl_unicode/normalizer_is_normalized_detects_nfc
// origin: languages/php/tests/php/test_intl_unicode.rs

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

if (!class_exists('Normalizer')) { echo 'skip'; } else {
    echo Normalizer::isNormalized('café', Normalizer::FORM_C) ? 'yes' : 'no';
}

__vybe_check(ob_get_clean(), "yes");
