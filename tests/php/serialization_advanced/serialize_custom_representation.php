<?php
// vybe-test: php/serialization_advanced/serialize_custom_representation
// origin: languages/php/tests/php/test_serialization_advanced.rs

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

class Vector {
    public function __construct(public float $x, public float $y, public float $z) {}
    public function __serialize(): array { return [$this->x, $this->y, $this->z]; }
    public function __unserialize(array $data): void {
        [$this->x, $this->y, $this->z] = $data;
    }
    public function length(): float { return sqrt($this->x**2 + $this->y**2 + $this->z**2); }
}
$v = new Vector(1.0, 0.0, 0.0);
$s = serialize($v);
$v2 = unserialize($s);
echo round($v2->length(), 4);

__vybe_check(ob_get_clean(), "1");
