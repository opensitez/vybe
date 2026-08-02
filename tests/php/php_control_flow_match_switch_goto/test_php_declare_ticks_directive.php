<?php
// vybe-test: php/php_control_flow_match_switch_goto/test_php_declare_ticks_directive
// origin: languages/php/tests/php/test_php_control_flow_match_switch_goto.rs
// vybe-test-mode: compile

declare(ticks=1);
$a = 1;
$b = 2;
echo $a + $b;
