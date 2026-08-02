<?php
// vybe-test: php/stream_resolve_include_path/stream_resolve_include_path_exists
// origin: languages/php/tests/php/test_stream_resolve_include_path.rs

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

// Since we can't reliably predict the include path, we'll just test if it returns a string or false
$path = stream_resolve_include_path("php://memory");
echo is_string($path) || is_bool($path) ? "ok" : "fail";

__vybe_check(ob_get_clean(), "ok");
