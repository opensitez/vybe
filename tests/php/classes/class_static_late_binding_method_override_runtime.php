<?php
// vybe-test: php/classes/class_static_late_binding_method_override_runtime
// origin: languages/php/tests/php/test_classes.rs

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
    public static function factory(): static {
        return new static();
    }
}
class Child extends Base {}
echo (new Base())::factory() instanceof Base ? 'base' : 'not';
echo '|';
echo Child::factory() instanceof Child ? 'child' : 'no';

__vybe_check(ob_get_clean(), "base|child");
