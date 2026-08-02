<?php
// vybe-test: php/interfaces_deep/interface_union_type_accepts_contracts_runtime
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

interface Readable { public function read(): string; }
interface Writable { public function write(string $value): void; }

class Buffer implements Readable, Writable {
    private string $value = '';
    public function read(): string { return $this->value; }
    public function write(string $value): void { $this->value = $value; }
}

function appendSuffix(Readable|Writable $target): string {
    if ($target instanceof Writable) {
        $target->write('ok');
    }
    if ($target instanceof Readable) {
        return $target->read();
    }
    return '';
}
$target = new Buffer();
echo appendSuffix($target);
// output reflects the write done above
echo '|' . $target->read();

__vybe_check(ob_get_clean(), "ok|ok");
