<?php
// vybe-test: php/arrayaccess_countable/arrayaccess_offset_exists_and_get
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

class Box implements ArrayAccess {
    private array $data = [];
    public function offsetExists(mixed $k): bool { return isset($this->data[$k]); }
    public function offsetGet(mixed $k): mixed { return $this->data[$k]; }
    public function offsetSet(mixed $k, mixed $v): void { $this->data[$k] = $v; }
    public function offsetUnset(mixed $k): void { unset($this->data[$k]); }
}
$b = new Box;
$b['x'] = 10;
echo $b['x'];

__vybe_check(ob_get_clean(), "10");
