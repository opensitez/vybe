<?php
// vybe-test: php/match_advanced/match_stops_after_first_matching_condition_in_true_subject
// origin: languages/php/tests/php/test_match_advanced.rs

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

echo "match_stops_after_first_matching_condition_in_true_subject_ok";

__vybe_check(ob_get_clean(), "match_stops_after_first_matching_condition_in_true_subject_ok");
