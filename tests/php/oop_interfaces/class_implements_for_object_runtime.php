<?php
// vybe-test: php/oop_interfaces/class_implements_for_object_runtime
// origin: languages/php/tests/php/test_oop_interfaces.rs

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

interface Reader {
    public function read(): string;
}
interface Writer {
    public function write(string $v): void;
}
class Logger implements Reader, Writer {
    public function read(): string { return 'r'; }
    public function write(string $v): void { $this->value = $v; }
    private string $value = '';
    public function value(): string { return $this->value; }
}
$logger = new Logger();
$impl = class_implements($logger);
echo (isset($impl[Reader::class]) ? 'R' : '?') . (isset($impl[Writer::class]) ? 'W' : '?');

__vybe_check(ob_get_clean(), "RW");
