<?php
// vybe-test: php/php_streams_file_system_wrapper_ops/test_php_stream_wrapper_register_custom
// origin: languages/php/tests/php/test_php_streams_file_system_wrapper_ops.rs
// vybe-test-mode: compile

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
