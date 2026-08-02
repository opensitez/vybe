<?php
// vybe-test: php/php_oop_inheritance_abstract_final/test_php_self_static_class_binding_runtime
// origin: languages/php/tests/php/test_php_oop_inheritance_abstract_final.rs

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

abstract class BaseType {
    public static function selfClass(): string { return self::class; }
    public static function staticClass(): string { return static::class; }
}
class ChildType extends BaseType {}
echo BaseType::selfClass();
echo '|';
echo BaseType::staticClass();
echo '|';
echo ChildType::selfClass();
echo '|';
echo ChildType::staticClass();

__vybe_check(ob_get_clean(), "BaseType|BaseType|BaseType|ChildType");
