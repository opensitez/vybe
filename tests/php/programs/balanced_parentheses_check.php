<?php
// vybe-test: php/programs/balanced_parentheses_check
// origin: languages/php/tests/php/test_programs.rs

function __vybe_check($got, $want) {
    // Match the Rust harness's normalisation: strip \r, then drop trailing
    // newlines (it split on "\n" and popped empty trailing elements).
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    // Replay the program's own output so running the file by hand still
    // behaves like the program it was extracted from.
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

function isBalanced(string $s): bool {
    $stack = [];
    $pairs = [')'=>'(', ']'=>'[', '}'=>'{'];
    foreach (str_split($s) as $c) {
        if (in_array($c, ['(','[','{'])) $stack[] = $c;
        elseif (isset($pairs[$c])) {
            if (empty($stack) || array_pop($stack) !== $pairs[$c]) return false;
        }
    }
    return empty($stack);
}
echo isBalanced('{[()]}') ? 'true' : 'false';
echo "\n";
echo isBalanced('([)]') ? 'true' : 'false';
echo "\n";
echo isBalanced('((())') ? 'true' : 'false';
echo "\n";

__vybe_check(ob_get_clean(), "true\nfalse\nfalse");
