<?php
// vybe-test: php/stream_context_set_default/stream_context_set_default_opts
// origin: languages/php/tests/php/test_stream_context_set_default.rs

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

$opts = ['http' => ['method' => 'POST']];
$ctx = stream_context_set_default($opts);
$params = stream_context_get_options(stream_context_get_default());
echo $params['http']['method'];

__vybe_check(ob_get_clean(), "POST");
