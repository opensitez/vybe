<?php
// vybe-test: php/php_callables_is_callable_call_user_func/test_php_invokable_object_call_user_func
// origin: languages/php/tests/php/test_php_callables_is_callable_call_user_func.rs

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

class InvokableTransformer {
    public function __invoke(string $text): string {
        return str_rot13($text);
    }
}

$transformer = new InvokableTransformer();
echo call_user_func($transformer, "Hello");

__vybe_check(ob_get_clean(), "Uryyb");
