<?php
// vybe-test: php/string_regex_extended/preg_match_all_capture_groups
// origin: languages/php/tests/php/test_string_regex_extended.rs

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

preg_match_all('/(\w+)=(\w+)/', 'a=1 b=2 c=3', $m);
echo implode(',', $m[1]) . ':' . implode(',', $m[2]);
echo "\n";

__vybe_check(ob_get_clean(), "a,b,c:1,2,3");
