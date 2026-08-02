<?php
// vybe-test: php/debug_backtrace/backtrace_provides_file_key_in_frame
// origin: languages/php/tests/php/test_debug_backtrace.rs

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

function hasFile(): string {
    $f = debug_backtrace()[0]['file'] ?? '';
    return $f !== '' ? 'file' : 'none';
}
echo hasFile();

__vybe_check(ob_get_clean(), "file");
