<?php
// vybe-test: php/php_stream_wrapper_register_custom_protocol/test_php_stream_wrapper_stat_filesize
// origin: languages/php/tests/php/test_php_stream_wrapper_register_custom_protocol.rs
// vybe-test-mode: compile

class StatWrapper {
    public function url_stat(string $path, int $flags): array {
        return ["size" => 1024, 7 => 1024];
    }
}
stream_wrapper_register("statproto", StatWrapper::class);
$size = filesize("statproto://virtual");
stream_wrapper_unregister("statproto");
echo $size === 1024 ? "URL_STAT_OK" : "FAIL";
