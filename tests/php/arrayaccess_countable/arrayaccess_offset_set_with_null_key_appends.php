<?php
// vybe-test: php/arrayaccess_countable/arrayaccess_offset_set_with_null_key_appends
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

class List3 implements ArrayAccess {
    private array $d = [];
    public function offsetExists(mixed $k): bool { return isset($this->d[$k]); }
    public function offsetGet(mixed $k): mixed { return $this->d[$k]; }
    public function offsetSet(mixed $k, mixed $v): void {
        if ($k === null) $this->d[] = $v;
        else $this->d[$k] = $v;
    }
    public function offsetUnset(mixed $k): void { unset($this->d[$k]); }
    public function toArray(): array { return $this->d; }
}
$l = new List3;
$l[0] = 'head'; $l[] = 'tail'; $l[2] = 'mid';
echo implode('|', $l->toArray());

__vybe_check(ob_get_clean(), "head|tail|mid");
