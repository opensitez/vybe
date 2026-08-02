<?php
// vybe-test: php/modern_php_deep/fiber_scheduler_round_robin
// origin: languages/php/tests/php/test_modern_php_deep.rs

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

$fibers = [];
for ($i = 1; $i <= 3; $i++) {
    $n = $i;
    $fibers[] = new Fiber(function() use ($n) {
        echo "task$n start";
        Fiber::suspend();
        echo "task$n end";
    });
}
foreach ($fibers as $f) { $f->start(); }
foreach ($fibers as $f) { $f->resume(); }

__vybe_check(ob_get_clean(), "task1 starttask2 starttask3 starttask1 endtask2 endtask3 end");
