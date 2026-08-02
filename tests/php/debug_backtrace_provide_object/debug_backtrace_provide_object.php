<?php
// vybe-test: php/debug_backtrace_provide_object/debug_backtrace_provide_object
// origin: languages/php/tests/php/test_debug_backtrace_provide_object.rs

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

class TestClass {
    public function doTrace() {
        $trace = debug_backtrace(DEBUG_BACKTRACE_PROVIDE_OBJECT, 1);
        echo isset($trace[0]['object']) && $trace[0]['object'] instanceof TestClass ? "ok" : "fail";
    }
}
$t = new TestClass();
$t->doTrace();

__vybe_check(ob_get_clean(), "ok");
