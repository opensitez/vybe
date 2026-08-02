<?php
// vybe-test: php/php_control_flow_match_switch_goto/test_php_switch_loose_equality_vs_match_strict
// origin: languages/php/tests/php/test_php_control_flow_match_switch_goto.rs
// vybe-test-mode: compile

$val = "0";
$switchRes = "";
switch ($val) {
    case 0: $switchRes = "LOOSE_MATCH"; break;
    case "0": $switchRes = "STRICT_MATCH"; break;
}
echo $switchRes;
