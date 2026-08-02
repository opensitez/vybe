<?php
// vybe-test: php/php_dynamic_calling/php_dynamic_calling_variable_function_alias_chain
// origin: languages/php/tests/php/test_php_dynamic_calling.rs

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

function source(): string { return 'base'; }
function transform(string $s): string { return strtoupper($s); }
$step1 = 'source';
$step2 = 'transform';
echo $step2($step1());

__vybe_check(ob_get_clean(), "BASE");
