<?php
// vybe-test: php/oop_advanced/dynamic_method_call_on_object
// origin: languages/php/tests/php/test_oop_advanced.rs

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

class Worker {
    public function run(string $task): string { return "run:$task"; }
    public function stop(): string { return "stop"; }
}
$w = new Worker();
$method = "run";
echo $w->$method("build");
echo "|";
echo $w->{"stop"}();

__vybe_check(ob_get_clean(), "run:build|stop");
