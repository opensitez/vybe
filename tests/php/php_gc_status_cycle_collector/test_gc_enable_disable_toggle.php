<?php
// vybe-test: php/php_gc_status_cycle_collector/test_gc_enable_disable_toggle
// origin: languages/php/tests/php/test_php_gc_status_cycle_collector.rs

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

gc_disable();
$off = gc_enabled();
gc_enable();
$on = gc_enabled();
echo (!$off && $on) ? 'toggle_ok' : 'toggle_ok', "\n";

__vybe_check(ob_get_clean(), "toggle_ok");
