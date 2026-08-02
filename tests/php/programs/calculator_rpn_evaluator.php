<?php
// vybe-test: php/programs/calculator_rpn_evaluator
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

function rpn(string $expr): float {
    $stack = [];
    foreach (explode(' ', $expr) as $token) {
        if (is_numeric($token)) {
            $stack[] = (float)$token;
        } else {
            $b = array_pop($stack);
            $a = array_pop($stack);
            match($token) {
                '+' => $stack[] = $a + $b,
                '-' => $stack[] = $a - $b,
                '*' => $stack[] = $a * $b,
                '/' => $stack[] = $a / $b,
            };
        }
    }
    return array_pop($stack);
}
echo rpn('3 4 +') . "\n";
echo rpn('5 1 2 + 4 * + 3 -') . "\n";

__vybe_check(ob_get_clean(), "7\n14");
