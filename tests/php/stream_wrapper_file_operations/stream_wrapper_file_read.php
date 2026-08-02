<?php
// vybe-test: php/stream_wrapper_file_operations/stream_wrapper_file_read
// origin: languages/php/tests/php/test_stream_wrapper_file_operations.rs

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

class ReadWrapper {
    private $position = 0;
    private $data = "hello stream";
    
    public function stream_open($path, $mode, $options, &$opened_path) {
        return true;
    }
    
    public function stream_read($count) {
        $ret = substr($this->data, $this->position, $count);
        $this->position += strlen($ret);
        return $ret;
    }
    
    public function stream_eof() {
        return $this->position >= strlen($this->data);
    }
    
    public function stream_stat() {
        return [];
    }
}
stream_wrapper_register("readproto", "ReadWrapper");
$fp = fopen("readproto://test", "r");
echo fread($fp, 5);
echo fread($fp, 7);
fclose($fp);

__vybe_check(ob_get_clean(), "hello stream");
