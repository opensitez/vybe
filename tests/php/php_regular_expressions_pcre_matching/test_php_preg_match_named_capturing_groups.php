<?php
// vybe-test: php/php_regular_expressions_pcre_matching/test_php_preg_match_named_capturing_groups
// origin: languages/php/tests/php/test_php_regular_expressions_pcre_matching.rs

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

$pattern = '/(?P<year>\d{4})-(?P<month>\d{2})-(?P<day>\d{2})/';
$subject = "Date: 2024-05-12";
if (preg_match($pattern, $subject, $matches)) {
    echo "Year={$matches['year']} Month={$matches['month']} Day={$matches['day']}";
}

__vybe_check(ob_get_clean(), "Year=2024 Month=05 Day=12");
