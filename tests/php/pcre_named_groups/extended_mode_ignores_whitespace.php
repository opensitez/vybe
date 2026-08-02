<?php
// vybe-test: php/pcre_named_groups/extended_mode_ignores_whitespace
// origin: languages/php/tests/php/test_pcre_named_groups.rs

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

$pattern = '/
    (\d{4})  # year
    -
    (\d{2})  # month
/x';
preg_match($pattern, '2024-03', $m);
echo $m[1] . '-' . $m[2];

__vybe_check(ob_get_clean(), "2024-03");
