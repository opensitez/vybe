<?php
// vybe-test: php/phase2/attribute_on_function
// origin: languages/php/tests/php/test_phase2.rs
// vybe-test-mode: compile

#[Pure] function add(int $a, int $b): int { return $a + $b; }
