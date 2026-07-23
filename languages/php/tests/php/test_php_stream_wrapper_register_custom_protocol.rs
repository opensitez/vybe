use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Stream Wrappers: stream_wrapper_register & Custom Protocols
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_stream_wrapper_register_memory_protocol() {
    let out = run_prints(
        r##"<?php
class MemoryStreamWrapper {
    public $context;
    private string $data = "";
    private int $position = 0;

    public function stream_open(string $path, string $mode, int $options, ?string &$opened_path): bool {
        $this->position = 0;
        return true;
    }

    public function stream_write(string $data): int {
        $this->data .= $data;
        $this->position += strlen($data);
        return strlen($data);
    }

    public function stream_read(int $count): string {
        $ret = substr($this->data, $this->position, $count);
        $this->position += strlen($ret);
        return $ret;
    }

    public function stream_tell(): int { return $this->position; }
    public function stream_eof(): bool { return $this->position >= strlen($this->data); }
}

stream_wrapper_register("memoryvar", MemoryStreamWrapper::class);

$fp = fopen("memoryvar://test", "r+");
fwrite($fp, "Hello Custom Stream");
rewind($fp);

$read = stream_get_contents($fp);
fclose($fp);
stream_wrapper_unregister("memoryvar");

echo $read;
"##,
    );
    assert_eq!(out, vec!["Hello Custom Stream"]);
}

#[test]
fn test_php_stream_get_wrappers_includes_registered() {
    let out = run_prints(
        r##"<?php
class DummyWrapper { public function stream_open(): bool { return true; } }
stream_wrapper_register("dummyproto", DummyWrapper::class);

$wrappers = stream_get_wrappers();
$hasDummy = in_array("dummyproto", $wrappers);
stream_wrapper_unregister("dummyproto");

echo $hasDummy ? "WRAPPER_REGISTERED" : "FAIL";
"##,
    );
    assert_eq!(out, vec!["WRAPPER_REGISTERED"]);
}

#[test]
fn test_php_stream_wrapper_restore_builtin_protocol() {
    compile_ok(
        r##"<?php
stream_wrapper_unregister("file");
echo !in_array("file", stream_get_wrappers()) ? "UNREGISTERED_FILE" : "FAIL";
stream_wrapper_restore("file");
echo in_array("file", stream_get_wrappers()) ? " RESTORED_FILE" : " FAIL";
"##,
    );
}

#[test]
fn test_php_stream_wrapper_stat_filesize() {
    compile_ok(
        r##"<?php
class StatWrapper {
    public function url_stat(string $path, int $flags): array {
        return ["size" => 1024, 7 => 1024];
    }
}
stream_wrapper_register("statproto", StatWrapper::class);
$size = filesize("statproto://virtual");
stream_wrapper_unregister("statproto");
echo $size === 1024 ? "URL_STAT_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_stream_wrapper_unlink_support() {
    compile_ok(
        r##"<?php
class UnlinkWrapper {
    public function unlink(string $path): bool { return true; }
}
stream_wrapper_register("unlinkproto", UnlinkWrapper::class);
$res = unlink("unlinkproto://dummy");
stream_wrapper_unregister("unlinkproto");
echo $res ? "UNLINK_WRAPPER_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_stream_wrapper_mkdir_rmdir() {
    compile_ok(
        r##"<?php
class DirWrapper {
    public function mkdir(string $path, int $mode, int $options): bool { return true; }
    public function rmdir(string $path, int $options): bool { return true; }
}
stream_wrapper_register("dirproto", DirWrapper::class);
$m = mkdir("dirproto://newdir");
$r = rmdir("dirproto://newdir");
stream_wrapper_unregister("dirproto");
echo $m && $r ? "DIR_WRAPPER_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_stream_wrapper_rename_support() {
    compile_ok(
        r##"<?php
class RenameWrapper {
    public function rename(string $path_from, string $path_to): bool { return true; }
}
stream_wrapper_register("renameproto", RenameWrapper::class);
$res = rename("renameproto://a", "renameproto://b");
stream_wrapper_unregister("renameproto");
echo $res ? "RENAME_WRAPPER_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_stream_wrapper_duplicate_register_error() {
    compile_ok(
        r##"<?php
class DupWrapper {}
stream_wrapper_register("dupproto", DupWrapper::class);
$res = @stream_wrapper_register("dupproto", DupWrapper::class);
stream_wrapper_unregister("dupproto");
echo $res === false ? "DUPLICATE_REGISTER_FALSE" : "FAIL";
"##,
    );
}

#[test]
fn test_php_stream_wrapper_dir_opendir_readdir() {
    compile_ok(
        r##"<?php
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
"##,
    );
}

#[test]
fn test_php_stream_wrapper_stream_seek_tell() {
    compile_ok(
        r##"<?php
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
"##,
    );
}
