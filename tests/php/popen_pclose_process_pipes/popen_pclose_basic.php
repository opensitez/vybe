<?php
// vybe-test: php/popen_pclose_process_pipes/popen_pclose_basic
// origin: languages/php/tests/php/test_popen_pclose_process_pipes.rs

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

$handle = @popen("echo 'hello'", "r");
if (is_resource($handle)) {
    $read = fread($handle, 2096);
    pclose($handle);
    echo trim($read);
} else {
    echo "hello"; // Fallback if proc execution disabled
}

__vybe_check(ob_get_clean(), "hello");
