<?php
// vybe-test: php/arrayaccess_countable/arrayaccess_offset_get_missing_key_returns_null
// origin: languages/php/tests/php/test_arrayaccess_countable.rs

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

class BagMiss implements ArrayAccess {
    public function offsetExists(mixed $k): bool { return false; }
    public function offsetGet(mixed $k): mixed { return null; }
    public function offsetSet(mixed $k, mixed $v): void {}
    public function offsetUnset(mixed $k): void {}
}
$b = new BagMiss;
echo var_export($b['missing'], true);

__vybe_check(ob_get_clean(), "NULL");
