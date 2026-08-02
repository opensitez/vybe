<?php
// vybe-test: php/php_control_flow_match_switch_goto/test_php_goto_loop_restart_pattern
// origin: languages/php/tests/php/test_php_control_flow_match_switch_goto.rs
// vybe-test-mode: compile

$attempts = 0;
retry:
$attempts++;
if ($attempts < 3) {
    goto retry;
}
echo "Attempts: $attempts";
