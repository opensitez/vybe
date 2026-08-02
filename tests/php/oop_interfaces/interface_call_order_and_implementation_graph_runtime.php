<?php
// vybe-test: php/oop_interfaces/interface_call_order_and_implementation_graph_runtime
// origin: languages/php/tests/php/test_oop_interfaces.rs

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

interface Logger {
    public function emit(string $message): string;
}
interface Timestamped {
    public function emit(string $message): string;
}
class Audit implements Logger, Timestamped {
    public function emit(string $message): string {
        return 'log:' . $message;
    }
}
echo (new Audit)->emit('x');

__vybe_check(ob_get_clean(), "log:x");
