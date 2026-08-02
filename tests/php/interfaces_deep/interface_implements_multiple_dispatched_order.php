<?php
// vybe-test: php/interfaces_deep/interface_implements_multiple_dispatched_order
// origin: languages/php/tests/php/test_interfaces_deep.rs

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

interface Reader { public function read(): string; }
interface Writer { public function write(string $v): string; }

class Document implements Reader, Writer {
    public function read(): string { return 'read'; }
    public function write(string $v): string { return 'write:' . $v; }
}

$doc = new Document();
echo $doc->read() . '|' . $doc->write('v');

__vybe_check(ob_get_clean(), "read|write:v");
