<?php
// vybe-test: php/php_stream_wrapper_register_custom_protocol/test_php_stream_wrapper_stream_seek_tell
// origin: languages/php/tests/php/test_php_stream_wrapper_register_custom_protocol.rs
// vybe-test-mode: compile

class SeekWrapper {
    private int $pos = 0;
    public function stream_open(): bool { return true; }
    public function stream_seek(int $offset, int $whence): bool { $this->pos = $offset; return true; }
    public function stream_tell(): int { return $this->pos; }
}
stream_wrapper_register("seekproto", SeekWrapper::class);
$fp = fopen("seekproto://file", "r");
fseek($fp, 50);
$tell = ftell($fp);
fclose($fp);
stream_wrapper_unregister("seekproto");
echo $tell === 50 ? "SEEK_TELL_WRAPPER_OK" : "FAIL";
