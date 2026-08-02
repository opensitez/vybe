<?php
// vybe-test: php/stream_wrapper_dir_operations/stream_wrapper_dir_rewinddir
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

class RewindDirWrapper {
    private $files = ['a', 'b'];
    private $index = 0;
    
    public function dir_opendir($path, $options) {
        return true;
    }
    
    public function dir_readdir() {
        if ($this->index < count($this->files)) {
            return $this->files[$this->index++];
        }
        return false;
    }
    
    public function dir_rewinddir() {
        $this->index = 0;
        return true;
    }
}
stream_wrapper_register("rewindproto", "RewindDirWrapper");
$dir = opendir("rewindproto://mydir");
echo readdir($dir);
rewinddir($dir);
echo readdir($dir);
closedir($dir);

__vybe_check(ob_get_clean(), "aa");
