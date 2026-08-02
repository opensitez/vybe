<?php
// vybe-test: php/scope_patterns/global_keyword_modify
// origin: languages/php/tests/php/test_scope_patterns.rs
// vybe-test-mode: compile

$total = 0;
function addToTotal(int $n): void {
    global $total;
    $total += $n;
}
addToTotal(5);
addToTotal(3);
echo $total;
