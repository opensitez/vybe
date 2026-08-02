<?php
// vybe-test: php/oop_advanced/variadic_override_in_child_class
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

class Base {
    public function combine(string ...$parts): string {
        return implode("-", $parts);
    }
}
class Child extends Base {
    public function combine(string ...$parts): string {
        $upper = array_map("strtoupper", $parts);
        return parent::combine(...$upper);
    }
}
$c = new Child();
echo $c->combine("foo", "bar", "baz"), "\n";

__vybe_check(ob_get_clean(), "FOO-BAR-BAZ");
