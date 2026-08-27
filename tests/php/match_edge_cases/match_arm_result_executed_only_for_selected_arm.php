<?php
// vybe-test: php/match_edge_cases/match_arm_result_executed_only_for_selected_arm
// origin: languages/php/tests/php/test_match_edge_cases.rs

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

echo "match_arm_result_executed_only_for_selected_arm_ok";

__vybe_check(ob_get_clean(), "match_arm_result_executed_only_for_selected_arm_ok");
