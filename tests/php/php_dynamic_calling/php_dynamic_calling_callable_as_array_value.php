<?php
// vybe-test: php/php_dynamic_calling/php_dynamic_calling_callable_as_array_value
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

$callables = [
    'twice' => function(int $v): int { return $v * 2; },
    'sum' => function(int $a, int $b): int { return $a + $b; },
];
echo $callables['twice'](4);
echo '|';
echo call_user_func_array($callables['sum'], [3, 5]);

__vybe_check(ob_get_clean(), "8|8");
