
crate::php_cases! {
    stream_wrapper_file_read => {
        r#"<?php
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
"#,
        ["hello stream"]
    };

    stream_wrapper_file_write => {
        r#"<?php
class WriteWrapper {
    public static $buffer = "";
    
    public function stream_open($path, $mode, $options, &$opened_path) {
        return true;
    }
    
    public function stream_write($data) {
        self::$buffer .= $data;
        return strlen($data);
    }
}
stream_wrapper_register("writeproto", "WriteWrapper");
$fp = fopen("writeproto://test", "w");
fwrite($fp, "part1");
fwrite($fp, "part2");
fclose($fp);
echo WriteWrapper::$buffer;
"#,
        ["part1part2"]
    };

    stream_wrapper_file_tell_seek => {
        r#"<?php
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
"#,
        ["1015"]
    };

    stream_wrapper_file_stat => {
        r#"<?php
class StatWrapper {
    public function stream_open($path, $mode, $options, &$opened_path) {
        return true;
    }
    
    public function stream_stat() {
        return [
            'size' => 1024,
            'mode' => 0100644,
        ];
    }
}
stream_wrapper_register("statproto", "StatWrapper");
$fp = fopen("statproto://test", "r");
$stat = fstat($fp);
echo $stat['size'] . '-' . decoct($stat['mode']);
fclose($fp);
"#,
        ["1024-100644"]
    };
}
