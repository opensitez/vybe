<?php
// vybe-test: php/fiber_lifecycle/fiber_inside_class_method
// origin: languages/php/tests/php/test_fiber_lifecycle.rs

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
    public function run(): Fiber {
        return new Fiber(function(): void {
            $result = Fiber::suspend('ready');
            echo "processed: $result";
        });
    }
}
$fiber = (new Worker())->run();
$fiber->start();
$fiber->resume('input');

__vybe_check(ob_get_clean(), "processed: input");
