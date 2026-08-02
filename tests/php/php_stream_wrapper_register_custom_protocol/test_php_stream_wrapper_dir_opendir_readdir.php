<?php
// vybe-test: php/php_stream_wrapper_register_custom_protocol/test_php_stream_wrapper_dir_opendir_readdir
// origin: languages/php/tests/php/test_php_stream_wrapper_register_custom_protocol.rs
// vybe-test-mode: compile

class OpenDirWrapper {
    private array $files = ["file1.txt", "file2.txt"];
    private int $idx = 0;
    public function dir_opendir(string $path, int $options): bool { $this->idx = 0; return true; }
    public function dir_readdir(): string|bool { return $this->files[$this->idx++] ?? false; }
}
stream_wrapper_register("opendirproto", OpenDirWrapper::class);
$dh = opendir("opendirproto://folder");
$f1 = readdir($dh);
closedir($dh);
stream_wrapper_unregister("opendirproto");
echo $f1 === "file1.txt" ? "READDIR_WRAPPER_OK" : "FAIL";
