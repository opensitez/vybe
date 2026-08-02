<?php
// vybe-test: php/abstract_final_patterns/final_class_late_static_binding_stays_in_class
// origin: languages/php/tests/php/test_abstract_final_patterns.rs

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

final class Singleton {
    private static ?self $instance = null;
    private function __construct() {}
    public static function get(): static { return self::$instance ??= new self(); }
    public function whoAmI(): string { return static::class; }
}
echo Singleton::get()->whoAmI(), "\n";

__vybe_check(ob_get_clean(), "Singleton");
