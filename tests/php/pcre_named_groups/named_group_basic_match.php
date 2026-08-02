<?php
// vybe-test: php/pcre_named_groups/named_group_basic_match
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

preg_match('/(?P<year>\d{4})-(?P<month>\d{2})-(?P<day>\d{2})/', '2024-03-15', $m);
echo $m['year'] . ',' . $m['month'] . ',' . $m['day'];

__vybe_check(ob_get_clean(), "2024,03,15");
