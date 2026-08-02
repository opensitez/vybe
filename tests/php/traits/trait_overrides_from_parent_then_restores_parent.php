<?php
// vybe-test: php/traits/trait_overrides_from_parent_then_restores_parent
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

class Base {
    public function label(): string { return 'base'; }
}
trait T {
    public function label(): string { return 'trait'; }
}
class Child extends Base {
    use T {
        T::label as parentLabel;
    }
    public function local(): string { return $this->parentLabel(); }
}
$obj = new Child();
echo $obj->label() . '|' . $obj->local();

__vybe_check(ob_get_clean(), "trait|base");
