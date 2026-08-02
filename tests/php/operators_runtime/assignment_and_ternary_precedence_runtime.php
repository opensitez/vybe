<?php
// vybe-test: php/operators_runtime/assignment_and_ternary_precedence_runtime
// origin: languages/php/tests/php/test_operators_runtime.rs

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

$value = 0;
$result = $value ? 'truthy' : 'falsey';
echo $result;
echo '|';
$value = 1;
$result = $value > 0 ? 'gt0' : 'le0';
echo $result;
echo '|';
$value = 0;
echo $value ? 'first' : $value ? 'second' : 'third';
echo '|';
echo ($value ? 'first' : ($value ? 'second' : 'third'));

__vybe_check(ob_get_clean(), "falsey|gt0|third|third");
