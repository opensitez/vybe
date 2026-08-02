<?php
// vybe-test: php/patterns/proxy_lazy_loading
// origin: languages/php/tests/php/test_patterns.rs

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

interface ImageInterface {
    public function display(): string;
}
class RealImage implements ImageInterface {
    private $filename;
    public function __construct(string $f) {
        $this->filename = $f;
        echo 'loaded:' . $f;
    }
    public function display(): string { return 'showing:' . $this->filename; }
}
class ImageProxy implements ImageInterface {
    private $filename;
    private $real = null;
    public function __construct(string $f) { $this->filename = $f; }
    public function display(): string {
        if ($this->real === null) {
            $this->real = new RealImage($this->filename);
        }
        return $this->real->display();
    }
}
$img = new ImageProxy('photo.jpg');
echo 'proxy created';
echo $img->display();

__vybe_check(ob_get_clean(), "proxy createdloaded:photo.jpgshowing:photo.jpg");
