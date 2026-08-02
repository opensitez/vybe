<?php
// vybe-test: php/php_dynamic_calling/php_dynamic_calling_class_string_and_method_string_with_call_user_func
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

class DynamicMethodCarrier {
    public static function fromStatic(string $value): string { return 'static-' . $value; }
}
$class = DynamicMethodCarrier::class;
$method = 'fromStatic';
echo call_user_func([$class, $method], 'ok');

__vybe_check(ob_get_clean(), "static-ok");
