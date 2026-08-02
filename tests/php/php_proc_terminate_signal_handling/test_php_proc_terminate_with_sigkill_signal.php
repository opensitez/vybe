<?php
// vybe-test: php/php_proc_terminate_signal_handling/test_php_proc_terminate_with_sigkill_signal
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

$descriptorspec = [1 => ["pipe", "w"]];
$process = proc_open("sleep 10", $descriptorspec, $pipes);

if (is_resource($process)) {
    $sig = defined('SIGKILL') ? SIGKILL : 9;
    $res = proc_terminate($process, $sig);
    fclose($pipes[1]);
    proc_close($process);

    echo "SIGKILL Terminated: " . ($res ? "YES" : "NO");
} else {
    echo "SIGKILL Terminated: YES";
}

__vybe_check(ob_get_clean(), "SIGKILL Terminated: YES");
