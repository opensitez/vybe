<?php
// vybe-test: php/php_streams_file_system_wrapper_ops/test_php_stream_wrapper_register_custom
// origin: languages/php/tests/php/test_php_streams_file_system_wrapper_ops.rs

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

class VariableStream {
    public static string $data = "";
    public function stream_open($path, $mode, $options, &$opened_path) { return true; }
    public function stream_write($data) { self::$data .= $data; return strlen($data); }
    public function stream_read($count) { return ""; }
}

stream_wrapper_register("var", VariableStream::class);
file_put_contents("var://buffer", "custom stream data");
echo VariableStream::$data;
stream_wrapper_unregister("var");


__vybe_check(ob_get_clean(), "custom stream data");
