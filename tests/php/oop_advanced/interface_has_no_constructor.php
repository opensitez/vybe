<?php
// vybe-test: php/oop_advanced/interface_has_no_constructor
// origin: languages/php/tests/php/test_oop_advanced.rs
// vybe-test-mode: compile

interface Shape {
    public function area(): float;
    public function perimeter(): float;
}
class Rect implements Shape {
    public function __construct(private float $w, private float $h) {}
    public function area(): float { return $this->w * $this->h; }
    public function perimeter(): float { return 2 * ($this->w + $this->h); }
}
$r = new Rect(3.0, 4.0);
echo $r->area(), "\n";
echo $r->perimeter(), "\n";
