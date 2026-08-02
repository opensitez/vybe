<?php
// vybe-test: php/php_reflection_class_constant_modifiers/test_reflection_class_constant_is_final
// origin: languages/php/tests/php/test_php_reflection_class_constant_modifiers.rs

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

class FinalConstDemo {
    final public const LOCKED = 'immutable';
}
$rc = new ReflectionClass(FinalConstDemo::class);
$c = $rc->getReflectionConstant('LOCKED');
echo $c->isFinal() ? 'final_const' : 'overrideable', "\n";

__vybe_check(ob_get_clean(), "final_const");
