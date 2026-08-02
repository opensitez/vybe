<?php
// vybe-test: php/oop_interfaces/class_implements_multiple
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

interface Drawable { public function draw(): string; }
interface Resizable { public function resize(float $f): static; }
class Shape implements Drawable, Resizable {
    public function __construct(private float $size = 1.0) {}
    public function draw(): string { return "shape({$this->size})"; }
    public function resize(float $f): static { return new static($this->size * $f); }
}
$s = (new Shape(2.0))->resize(3.0);
echo $s->draw();

__vybe_check(ob_get_clean(), "shape(6)");
