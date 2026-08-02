<?php
// vybe-test: php/operators/comparison_operators_short_circuit_side_effects_runtime
// origin: languages/php/tests/php/test_operators.rs

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

$log = [];
$right = function() use (&$log): bool {
    $log[] = 'right';
    return false;
};

if (false && $right()) {
    echo 'bad';
}
echo (count($log) === 0 ? 'no-right' : 'right-called');
echo '|';
$log = [];
if (true || $right()) {
    echo 'skip';
}
echo (count($log) === 0 ? 'no-right' : 'right-called');
echo '|';
echo ($right() && true) ? 'bad' : 'ok';
echo '|';
echo count($log);

__vybe_check(ob_get_clean(), "no-right|skipno-right|ok|1");
