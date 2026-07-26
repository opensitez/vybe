
crate::php_cases! {
    stream_wrapper_dir_opendir_readdir => {
        r#"<?php
class DirWrapper {
    private $files = ['file1.txt', 'file2.txt'];
    private $index = 0;
    
    public function dir_opendir($path, $options) {
        $this->index = 0;
        return true;
    }
    
    public function dir_readdir() {
        if ($this->index < count($this->files)) {
            return $this->files[$this->index++];
        }
        return false;
    }
    
    public function dir_closedir() {
        return true;
    }
}
stream_wrapper_register("dirproto", "DirWrapper");
$dir = opendir("dirproto://mydir");
while (($file = readdir($dir)) !== false) {
    echo $file . ',';
}
closedir($dir);
"#,
        ["file1.txt,file2.txt,"]
    };

    stream_wrapper_dir_rewinddir => {
        r#"<?php
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
"#,
        ["aa"]
    };

    stream_wrapper_mkdir_rmdir => {
        r#"<?php
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
"#,
        ["mkdir:mkdirproto://newdir,rmdir:mkdirproto://newdir"]
    };

    stream_wrapper_rename_unlink => {
        r#"<?php
class FsWrapper {
    public static $log = [];
    
    public function rename($path_from, $path_to) {
        self::$log[] = "rename:$path_from->$path_to";
        return true;
    }
    
    public function unlink($path) {
        self::$log[] = "unlink:$path";
        return true;
    }
}
stream_wrapper_register("fsproto", "FsWrapper");
rename("fsproto://old", "fsproto://new");
unlink("fsproto://file");
echo implode(',', FsWrapper::$log);
"#,
        ["rename:fsproto://old->fsproto://new,unlink:fsproto://file"]
    };
}
