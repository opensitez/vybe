<?php
// vybe-test: php/debug_backtrace/backtrace_type_is_object_for_instance_method
// origin: languages/php/tests/php/test_debug_backtrace.rs

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

class Tracer {
    public function mark(): string {
        return debug_backtrace()[0]['type'] ?? '?';
    }
}
echo (new Tracer())->mark();

__vybe_check(ob_get_clean(), "->");
