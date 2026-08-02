<?php
// vybe-test: php/intersection_types/dnf_type_union_of_intersection_and_class
// origin: languages/php/tests/php/test_intersection_types.rs

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
interface Writable { public function write(string $s): void; }
class Stream implements Readable, Writable {
    private string $buf = '';
    public function read(): string { return $this->buf; }
    public function write(string $s): void { $this->buf .= $s; }
}
class NullStream {
    public function write(string $s): void {}
}
function writeIfPossible((Readable&Writable)|NullStream $s, string $data): void {
    $s->write($data);
}
$s = new Stream();
writeIfPossible($s, "hello");
echo $s->read();

__vybe_check(ob_get_clean(), "hello");
