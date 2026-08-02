<?php
// vybe-test: php/traits/trait_instantiate_static_context
// origin: languages/php/tests/php/test_traits.rs

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

trait Maker {
    public static function create(int $value): static {
        return new static($value);
    }
}
class Item {
    use Maker;
    public function __construct(private int $v) {}
    public function value(): int { return $this->v; }
}
echo Item::create(7)->value();

__vybe_check(ob_get_clean(), "7");
