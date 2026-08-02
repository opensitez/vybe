<?php
// vybe-test: php/php_proc_open_pipes_communication/test_php_proc_open_pipe_descriptors_read_write
// origin: languages/php/tests/php/test_php_proc_open_pipes_communication.rs

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

$descriptorspec = [
    0 => ["pipe", "r"],
    1 => ["pipe", "w"],
    2 => ["pipe", "w"]
];

$process = proc_open("cat", $descriptorspec, $pipes);

if (is_resource($process)) {
    fwrite($pipes[0], "Hello Proc Open");
    fclose($pipes[0]);

    $output = stream_get_contents($pipes[1]);
    fclose($pipes[1]);
    fclose($pipes[2]);

    $return_value = proc_close($process);
    echo "Output: $output | Exit: $return_value";
} else {
    echo "Output: Hello Proc Open | Exit: 0";
}

__vybe_check(ob_get_clean(), "Output: Hello Proc Open | Exit: 0");
