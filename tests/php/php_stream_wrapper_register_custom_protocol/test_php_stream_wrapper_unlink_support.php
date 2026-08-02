<?php
// vybe-test: php/php_stream_wrapper_register_custom_protocol/test_php_stream_wrapper_unlink_support
// origin: languages/php/tests/php/test_php_stream_wrapper_register_custom_protocol.rs
// vybe-test-mode: compile

class UnlinkWrapper {
    public function unlink(string $path): bool { return true; }
}
stream_wrapper_register("unlinkproto", UnlinkWrapper::class);
$res = unlink("unlinkproto://dummy");
stream_wrapper_unregister("unlinkproto");
echo $res ? "UNLINK_WRAPPER_OK" : "FAIL";
