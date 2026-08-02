<?php
// vybe-test: php/php_stream_wrapper_register_custom_protocol/test_php_stream_wrapper_rename_support
// origin: languages/php/tests/php/test_php_stream_wrapper_register_custom_protocol.rs
// vybe-test-mode: compile

class RenameWrapper {
    public function rename(string $path_from, string $path_to): bool { return true; }
}
stream_wrapper_register("renameproto", RenameWrapper::class);
$res = rename("renameproto://a", "renameproto://b");
stream_wrapper_unregister("renameproto");
echo $res ? "RENAME_WRAPPER_OK" : "FAIL";
