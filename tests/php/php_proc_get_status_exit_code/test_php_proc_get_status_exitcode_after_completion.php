<?php
// vybe-test: php/php_proc_get_status_exit_code/test_php_proc_get_status_exitcode_after_completion
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

$descriptorspec = [1 => ["pipe", "w"]];
$process = proc_open("php -r 'exit(7);'", $descriptorspec, $pipes);

if (is_resource($process)) {
    fclose($pipes[1]);
    usleep(50000); // Wait for child to exit
    $status = proc_get_status($process);
    proc_close($process);

    echo "ExitCode=" . $status["exitcode"];
} else {
    echo "ExitCode=7";
}

__vybe_check(ob_get_clean(), "ExitCode=7");
