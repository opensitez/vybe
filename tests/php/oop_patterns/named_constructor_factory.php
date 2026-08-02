<?php
// vybe-test: php/oop_patterns/named_constructor_factory
// origin: languages/php/tests/php/test_oop_patterns.rs
// vybe-test-mode: compile

class Color {
    private function __construct(
        private int $r,
        private int $g,
        private int $b
    ) {}
    public static function fromRGB(int $r, int $g, int $b): self {
        return new self($r, $g, $b);
    }
    public static function fromHex(string $hex): self {
        $hex = ltrim($hex, '#');
        return new self(
            hexdec(substr($hex, 0, 2)),
            hexdec(substr($hex, 2, 2)),
            hexdec(substr($hex, 4, 2))
        );
    }
    public function toCSS(): string { return "rgb({$this->r},{$this->g},{$this->b})"; }
}
$red  = Color::fromRGB(255, 0, 0);
echo $red->toCSS();
