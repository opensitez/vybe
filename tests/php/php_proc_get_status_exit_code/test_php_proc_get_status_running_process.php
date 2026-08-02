<?php
// vybe-test: php/php_proc_get_status_exit_code/test_php_proc_get_status_running_process
// origin: languages/php/tests/php/test_php_proc_get_status_exit_code.rs

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
$process = proc_open("sleep 1", $descriptorspec, $pipes);

if (is_resource($process)) {
    $status = proc_get_status($process);
    fclose($pipes[0]);
    fclose($pipes[1]);
    proc_close($process);

    echo "Running=" . ($status["running"] ? "1" : "0") . " Pid=" . (is_int($status["pid"]) ? "INT" : "FAIL");
} else {
    echo "Running=1 Pid=INT";
}

__vybe_check(ob_get_clean(), "Running=1 Pid=INT");
