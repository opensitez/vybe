<?php
// vybe-test: php/string_case_multibyte/mb_str_split_longer_length_returns_single_chunk
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

echo count(mb_str_split('hey', 10)); echo '|'; echo mb_str_split('hey', 10)[0];

__vybe_check(ob_get_clean(), "1|hey");
