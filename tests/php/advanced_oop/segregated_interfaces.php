<?php
// vybe-test: php/advanced_oop/segregated_interfaces
// origin: languages/php/tests/php/test_advanced_oop.rs

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

interface Readable { public function read(): string; }
interface Writable { public function write(string $data): void; }
interface ReadWrite extends Readable, Writable {}
class Buffer implements ReadWrite {
    private string $buffer = '';
    public function read(): string { return $this->buffer; }
    public function write(string $data): void { $this->buffer .= $data; }
}
$b = new Buffer;
$b->write('hello');
$b->write(' world');
echo $b->read();

__vybe_check(ob_get_clean(), "hello world");
