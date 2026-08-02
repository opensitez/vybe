<?php
// vybe-test: php/php_dynamic_calling/php_dynamic_calling_callable_array_with_closure_and_string_target
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

function format_label(string $value): string {
    return 'L:' . $value;
}

$callables = [
    'format' => 'format_label',
    'twice' => function(string $value): string { return $value . $value; },
];

echo call_user_func($callables['format'], 'ok');
echo '|';
echo $callables['twice']('ha');

__vybe_check(ob_get_clean(), "L:ok|haha");
