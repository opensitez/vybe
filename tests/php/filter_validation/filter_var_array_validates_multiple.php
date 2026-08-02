<?php
// vybe-test: php/filter_validation/filter_var_array_validates_multiple
// origin: languages/php/tests/php/test_filter_validation.rs

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

$in = ['a' => '1', 'b' => 'x'];
$out = filter_var_array($in, ['a' => FILTER_VALIDATE_INT, 'b' => FILTER_VALIDATE_INT]);
echo ($out['a'] === 1 ? '1' : '0') . ($out['b'] === false ? 'f' : 't');

__vybe_check(ob_get_clean(), "1f");
