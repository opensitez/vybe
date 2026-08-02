<?php
// vybe-test: php/patterns/option_maybe_some_none
// origin: languages/php/tests/php/test_patterns.rs

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

class Option {
    private function __construct(private bool $hasValue, private $value = null) {}
    public static function some($v): self { return new self(true, $v); }
    public static function none(): self { return new self(false); }
    public function isSome(): bool { return $this->hasValue; }
    public function get() { return $this->value; }
    public function map(callable $fn): self {
        if (!$this->hasValue) return self::none();
        return self::some($fn($this->value));
    }
    public function getOrElse($default) { return $this->hasValue ? $this->value : $default; }
}
$opt = Option::some(10)->map(fn($x) => $x * 2);
echo $opt->isSome() ? 'some' : 'none';
echo $opt->get();
$empty = Option::none()->map(fn($x) => $x * 2);
echo $empty->getOrElse(99);

__vybe_check(ob_get_clean(), "some2099");
