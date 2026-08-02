<?php
// vybe-test: php/fiber_lifecycle/fiber_simulates_async_task_queue
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

$tasks = [];
$tasks[] = new Fiber(function(): void { echo "task1:start\n"; Fiber::suspend(); echo "task1:end\n"; });
$tasks[] = new Fiber(function(): void { echo "task2:start\n"; Fiber::suspend(); echo "task2:end\n"; });
foreach ($tasks as $t) $t->start();
foreach ($tasks as $t) if ($t->isSuspended()) $t->resume();

__vybe_check(ob_get_clean(), "task1:start\ntask2:start\ntask1:end\ntask2:end");
