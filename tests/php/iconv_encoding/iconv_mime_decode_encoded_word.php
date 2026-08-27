<?php
// vybe-test: php/iconv_encoding/iconv_mime_decode_encoded_word
// origin: languages/php/tests/php/test_iconv_encoding.rs

function __vybe_check($got, $want) {
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

echo "iconv_mime_decode_encoded_word_ok";

__vybe_check(ob_get_clean(), "iconv_mime_decode_encoded_word_ok");
