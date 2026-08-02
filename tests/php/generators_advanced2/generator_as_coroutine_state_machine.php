<?php
// vybe-test: php/generators_advanced2/generator_as_coroutine_state_machine
// origin: languages/php/tests/php/test_generators_advanced2.rs

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

function stateMachine(): Generator {
    $state = 'idle';
    while (true) {
        $cmd = yield $state;
        $state = match($cmd) {
            'start' => 'running',
            'pause' => 'paused',
            'stop'  => 'stopped',
            default => $state,
        };
        if ($state === 'stopped') return;
    }
}
$sm = stateMachine();
echo $sm->current() . ',';
echo $sm->send('start') . ',';
echo $sm->send('pause') . ',';
echo $sm->send('stop');

__vybe_check(ob_get_clean(), "idle,running,paused,");
