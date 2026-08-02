<?php
// vybe-test: php/oop_advanced/multiple_interfaces
// origin: languages/php/tests/php/test_oop_advanced.rs

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

interface Readable {
    public function read(): string;
}
interface Writable {
    public function write(string $data): void;
}
class File implements Readable, Writable {
    private string $content = "";
    public function read(): string { return $this->content; }
    public function write(string $data): void { $this->content .= $data; }
}
$f = new File();
$f->write("hello");
$f->write(" world");
echo $f->read(), "\n";

__vybe_check(ob_get_clean(), "hello world");
