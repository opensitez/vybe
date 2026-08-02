<?php
// vybe-test: php/stream_wrapper_file_operations/stream_wrapper_file_tell_seek
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

class SeekWrapper {
    private $position = 0;
    
    public function stream_open($path, $mode, $options, &$opened_path) {
        return true;
    }
    
    public function stream_tell() {
        return $this->position;
    }
    
    public function stream_seek($offset, $whence) {
        if ($whence === SEEK_SET) {
            $this->position = $offset;
        } elseif ($whence === SEEK_CUR) {
            $this->position += $offset;
        }
        return true;
    }
}
stream_wrapper_register("seekproto", "SeekWrapper");
$fp = fopen("seekproto://test", "r");
fseek($fp, 10, SEEK_SET);
echo ftell($fp);
fseek($fp, 5, SEEK_CUR);
echo ftell($fp);
fclose($fp);

__vybe_check(ob_get_clean(), "1015");
