<?php
// vybe-test: php/php_debug_backtrace_provide_object_option/test_debug_backtrace_provide_object
// origin: languages/php/tests/php/test_php_debug_backtrace_provide_object_option.rs

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

class TraceDemo {
    public string $id = 'demo_obj';
    public function traceSelf() {
        return debug_backtrace(DEBUG_BACKTRACE_PROVIDE_OBJECT, 1)[0];
    }
}
$td = new TraceDemo();
$frame = $td->traceSelf();
echo (isset($frame['object']) && $frame['object'] === $td) ? 'object_provided' : 'err', "\n";

__vybe_check(ob_get_clean(), "object_provided");
