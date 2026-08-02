<?php
// vybe-test: php/pcre_named_groups/preg_grep_inverted_filter
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

$nums = ['1', '2a', '3', '4b', '5'];
$notPure = preg_grep('/^\d+$/', $nums, PREG_GREP_INVERT);
echo implode(',', $notPure);

__vybe_check(ob_get_clean(), "2a,4b");
