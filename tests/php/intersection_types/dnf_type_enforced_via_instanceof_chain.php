<?php
// vybe-test: php/intersection_types/dnf_type_enforced_via_instanceof_chain
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

interface Readable3 { public function read(): string; }
interface Closeable2 { public function close(): void; }
class FileHandle implements Readable3, Closeable2 {
    private bool $closed = false;
    public function read(): string { return $this->closed ? '' : "data"; }
    public function close(): void { $this->closed = true; }
}
function readAndClose((Readable3&Closeable2)|null $handle): string {
    if ($handle === null) return 'null';
    $data = $handle->read();
    $handle->close();
    return $data;
}
echo readAndClose(new FileHandle()) . ',' . readAndClose(null);

__vybe_check(ob_get_clean(), "data,null");
