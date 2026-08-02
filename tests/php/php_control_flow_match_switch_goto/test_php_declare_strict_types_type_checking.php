<?php
// vybe-test: php/php_control_flow_match_switch_goto/test_php_declare_strict_types_type_checking
// origin: languages/php/tests/php/test_php_control_flow_match_switch_goto.rs
// vybe-test-mode: compile

declare(strict_types=1);

function addInts(int $a, int $b): int {
    return $a + $b;
}

echo addInts(10, 20);
