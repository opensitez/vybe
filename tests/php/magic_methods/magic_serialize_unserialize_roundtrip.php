<?php
// vybe-test: php/magic_methods/magic_serialize_unserialize_roundtrip
// origin: languages/php/tests/php/test_magic_methods.rs

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

class Vector2D {
    public function __construct(public float $x, public float $y) {}
    public function __serialize(): array { return ["x" => $this->x, "y" => $this->y]; }
    public function __unserialize(array $d): void { $this->x = $d["x"]; $this->y = $d["y"]; }
    public function length(): float { return sqrt($this->x ** 2 + $this->y ** 2); }
}
$v = new Vector2D(3.0, 4.0);
$raw = serialize($v);
$v2 = unserialize($raw);
echo $v2->x;
echo $v2->y;
echo $v2->length();

__vybe_check(ob_get_clean(), "345");
