<?php
// vybe-test: php/goto/goto_error_handling_pattern
// origin: languages/php/tests/php/test_goto.rs
// vybe-test-mode: compile

function riskyOp(int $n): string {
    if ($n < 0) goto error;
    if ($n === 0) goto zero;
    return "positive: $n";
    error:
    return "error: negative";
    zero:
    return "zero";
}
echo riskyOp(5);
echo riskyOp(0);
echo riskyOp(-1);
