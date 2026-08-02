<?php
// vybe-test: php/iconv_encoding/iconv_mime_encode_produces_encoded_word_prefix
// origin: languages/php/tests/php/test_iconv_encoding.rs

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

if (!function_exists('iconv_mime_encode')) { echo 'skip'; } else {
    $h = iconv_mime_encode('Subject', 'café', ['input-charset' => 'UTF-8', 'output-charset' => 'UTF-8']);
    echo str_starts_with($h, 'Subject: =?UTF-8?') ? 'mime' : 'raw';
}

__vybe_check(ob_get_clean(), "mime");
