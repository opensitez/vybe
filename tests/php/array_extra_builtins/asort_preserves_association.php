<?php
// vybe-test: php/array_extra_builtins/asort_preserves_association
// origin: languages/php/tests/php/test_array_extra_builtins.rs
// vybe-test-mode: compile

$stats = ["a" => 3, "b" => 1, "c" => 2];
asort($stats);
echo implode(",", array_keys($stats));
echo implode(",", $stats);
