<?php
// vybe-test: php/stream_wrapper_dir_operations/stream_wrapper_mkdir_rmdir
// origin: languages/php/tests/php/test_stream_wrapper_dir_operations.rs

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

class MkdirWrapper {
    public static $log = [];
    
    public function mkdir($path, $mode, $options) {
        self::$log[] = "mkdir:$path";
        return true;
    }
    
    public function rmdir($path, $options) {
        self::$log[] = "rmdir:$path";
        return true;
    }
}
stream_wrapper_register("mkdirproto", "MkdirWrapper");
mkdir("mkdirproto://newdir");
rmdir("mkdirproto://newdir");
echo implode(',', MkdirWrapper::$log);

__vybe_check(ob_get_clean(), "mkdir:mkdirproto://newdir,rmdir:mkdirproto://newdir");
