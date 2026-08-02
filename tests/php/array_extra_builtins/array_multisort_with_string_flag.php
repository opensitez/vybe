<?php
// vybe-test: php/array_extra_builtins/array_multisort_with_string_flag
// origin: languages/php/tests/php/test_array_extra_builtins.rs
// vybe-test-mode: compile

$scores = ["10", "2", "30"];
$names = ["low", "mid", "high"];
array_multisort($scores, SORT_NATURAL, SORT_ASC, $names, SORT_DESC);
echo implode(",", $scores);
echo implode(",", $names);
