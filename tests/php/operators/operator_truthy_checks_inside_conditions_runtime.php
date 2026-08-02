<?php
// vybe-test: php/operators/operator_truthy_checks_inside_conditions_runtime
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

function truthy_label(mixed $value): string {
    return $value ? 'T' : 'F';
}
$inputs = [null, 0, 1, '', '0', 'ok', [], [1], false, true];
$out = '';
	foreach ($inputs as $value) {
	    $out .= truthy_label($value);
	}
    echo $out;

__vybe_check(ob_get_clean(), "FFTFFTFTFT");
