<?php
// vybe-test: php/host_extra/stack_based_calculator
// origin: languages/php/tests/php/test_host_extra.rs
// vybe-test-mode: compile

$stack = new SplStack();
$tokens = explode(' ', '3 4 + 2 *');
foreach ($tokens as $token) {
    if (is_numeric($token)) {
        $stack->push(intval($token));
    } else {
        $b = $stack->pop();
        $a = $stack->pop();
        $result = match($token) {
            '+' => $a + $b,
            '-' => $a - $b,
            '*' => $a * $b,
            '/' => $a / $b,
            default => 0
        };
        $stack->push($result);
    }
}
echo $stack->pop();
