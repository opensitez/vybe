<?php
// vybe-test: php/magic_methods/magic_call_variadic_forwarding
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

class Wrapper {
    public function __construct(private object $target) {}
    public function __call(string $name, array $args) {
        if (method_exists($this->target, $name)) {
            return $this->target->$name(...$args);
        }
        return null;
    }
}
class Math {
    public function add(int $a, int $b): int { return $a + $b; }
    public function multiply(int $a, int $b, int $c): int { return $a * $b * $c; }
}
$w = new Wrapper(new Math());
echo $w->add(3, 4);
echo $w->multiply(2, 3, 4);

__vybe_check(ob_get_clean(), "724");
