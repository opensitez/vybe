<?php
// vybe-test: php/string_case_multibyte/mb_check_encoding_false_branch
// origin: languages/php/tests/php/test_string_case_multibyte.rs

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

$invalid = chr(255);
echo mb_check_encoding($invalid, 'UTF-8') ? 'ok' : 'bad';
echo '|';
echo mb_check_encoding('test', 'ISO-8859-1') ? 'lat' : 'no';

__vybe_check(ob_get_clean(), "bad|lat");
