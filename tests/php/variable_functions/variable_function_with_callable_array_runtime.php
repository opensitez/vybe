<?php
// vybe-test: php/variable_functions/variable_function_with_callable_array_runtime
// origin: languages/php/tests/php/test_variable_functions.rs

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

class Logger {
    public function label(string $name): string { return "log:$name"; }
}
$obj = new Logger();
$callable = [$obj, 'label'];
echo is_callable($callable) . '|';
echo $callable('trace');

__vybe_check(ob_get_clean(), "1|log:trace");
