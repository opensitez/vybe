<?php
// vybe-test: php/generators_advanced/generator_tracks_external_state
// origin: languages/php/tests/php/test_generators_advanced.rs

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

function makeTracker(array &$log) {
    return (function() use (&$log) {
        $log[] = "started";
        yield 1;
        $log[] = "middle";
        yield 2;
        $log[] = "ended";
    })();
}
$log = [];
$gen = makeTracker($log);
foreach ($gen as $v) {
    // consume
}
echo implode(",", $log);

__vybe_check(ob_get_clean(), "started,middle,ended");
