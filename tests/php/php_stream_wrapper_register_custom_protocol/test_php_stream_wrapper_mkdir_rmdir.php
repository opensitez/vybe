<?php
// vybe-test: php/php_stream_wrapper_register_custom_protocol/test_php_stream_wrapper_mkdir_rmdir
// origin: languages/php/tests/php/test_php_stream_wrapper_register_custom_protocol.rs
// vybe-test-mode: compile

class DirWrapper {
    public function mkdir(string $path, int $mode, int $options): bool { return true; }
    public function rmdir(string $path, int $options): bool { return true; }
}
stream_wrapper_register("dirproto", DirWrapper::class);
$m = mkdir("dirproto://newdir");
$r = rmdir("dirproto://newdir");
stream_wrapper_unregister("dirproto");
echo $m && $r ? "DIR_WRAPPER_OK" : "FAIL";
