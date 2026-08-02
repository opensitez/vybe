<?php
// vybe-test: php/magic_methods/magic_set_state_compile_structure
// origin: languages/php/tests/php/test_magic_methods.rs
// vybe-test-mode: compile

class Rectangle {
    public float $width;
    public float $height;
    public function __construct(float $w, float $h) {
        $this->width  = $w;
        $this->height = $h;
    }
    public static function __set_state(array $props): static {
        return new static($props['width'], $props['height']);
    }
    public function area(): float { return $this->width * $this->height; }
}
$r = Rectangle::__set_state(['width' => 4.0, 'height' => 5.0]);
echo $r->area();
