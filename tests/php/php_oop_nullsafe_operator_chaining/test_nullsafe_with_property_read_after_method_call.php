<?php
// vybe-test: php/php_oop_nullsafe_operator_chaining/test_nullsafe_with_property_read_after_method_call
// origin: languages/php/tests/php/test_php_oop_nullsafe_operator_chaining.rs

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

class Child {
    public function __construct(public ?Grandchild $nested = null) {}
    public function child(): ?Grandchild { return $this->nested; }
}
class Grandchild {
    public string $value = 'nested';
}
class Parent {
    public Child $child;
    public function __construct(?Child $child = null) { $this->child = $child ?? new Child(); }
}

echo (new Parent(new Child(new Grandchild())))->child()->child?->value ?? 'none';
echo '|';
echo (new Parent(new Child()))->child()->child?->value ?? 'none';

__vybe_check(ob_get_clean(), "nested|none");
