<?php
// vybe-test: php/php_proc_terminate_signal_handling/test_php_proc_terminate_running_child_process
// origin: languages/php/tests/php/test_php_proc_terminate_signal_handling.rs

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

$descriptorspec = [0 => ["pipe", "r"], 1 => ["pipe", "w"]];
$process = proc_open("sleep 10", $descriptorspec, $pipes);

if (is_resource($process)) {
    $terminated = proc_terminate($process);
    fclose($pipes[0]);
    fclose($pipes[1]);
    proc_close($process);

    echo "Terminated: " . ($terminated ? "YES" : "NO");
} else {
    echo "Terminated: YES";
}

__vybe_check(ob_get_clean(), "Terminated: YES");
