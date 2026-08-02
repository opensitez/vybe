<?php
// vybe-test: php/variable_functions/global_var_modified_in_function
// origin: languages/php/tests/php/test_variable_functions.rs
// vybe-test-mode: compile

$score = 0;
function addPoints(int $pts): void {
    global $score;
    $score += $pts;
}
addPoints(10);
addPoints(5);
echo $score;
