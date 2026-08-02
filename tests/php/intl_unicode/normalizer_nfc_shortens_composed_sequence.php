<?php
// vybe-test: php/intl_unicode/normalizer_nfc_shortens_composed_sequence
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
    $d = Normalizer::normalize("e\u{0301}", Normalizer::FORM_C);
    echo strlen($d) < strlen("e\u{0301}") ? 'nfc' : 'same';
}

__vybe_check(ob_get_clean(), "nfc");
