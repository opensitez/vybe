<?php
// vybe-test: php/edge_cases/multiple_constructors_via_static
// origin: languages/php/tests/php/test_edge_cases.rs
// vybe-test-mode: compile

class Color {
    public $r; public $g; public $b;
    public function __construct($r, $g, $b) { $this->r = $r; $this->g = $g; $this->b = $b; }
    public static function red() { return new Color(255, 0, 0); }
    public static function fromHex($hex) { return new Color(0, 0, 0); }
}
$red = Color::red();
echo $red->r;
