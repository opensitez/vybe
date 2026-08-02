<?php
// vybe-test: php/php_callables_is_callable_call_user_func/test_php_dynamic_method_name_invocation
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

class ActionHandler {
    public function handleInit(): string { return "INIT_DONE"; }
    public function handleRun(): string { return "RUN_DONE"; }
}

$handler = new ActionHandler();
$action = "Run";
$method = "handle" . $action;

echo $handler->$method();

__vybe_check(ob_get_clean(), "RUN_DONE");
