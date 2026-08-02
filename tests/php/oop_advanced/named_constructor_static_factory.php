<?php
// vybe-test: php/oop_advanced/named_constructor_static_factory
// origin: languages/php/tests/php/test_oop_advanced.rs

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

class Color {
    private function __construct(
        private int $r,
        private int $g,
        private int $b,
    ) {}
    public static function fromHex(string $hex): self {
        $hex = ltrim($hex, '#');
        return new self(
            hexdec(substr($hex, 0, 2)),
            hexdec(substr($hex, 2, 2)),
            hexdec(substr($hex, 4, 2)),
        );
    }
    public static function fromRgb(int $r, int $g, int $b): self {
        return new self($r, $g, $b);
    }
    public function __toString(): string {
        return "rgb({$this->r},{$this->g},{$this->b})";
    }
}
$c1 = Color::fromHex('#ff8000');
$c2 = Color::fromRgb(0, 128, 255);
echo $c1, "\n";
echo $c2, "\n";

__vybe_check(ob_get_clean(), "rgb(255,128,0)\nrgb(0,128,255)");
