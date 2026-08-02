<?php
// vybe-test: php/traits_advanced/trait_static_method
// origin: languages/php/tests/php/test_traits_advanced.rs

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

trait Singleton {
    private static ?self $instance = null;
    public static function getInstance(): static {
        if (static::$instance === null) static::$instance = new static();
        return static::$instance;
    }
}
class Config { use Singleton; public int $value = 42; }
$a = Config::getInstance(); $a->value = 99;
$b = Config::getInstance();
echo $b->value;

__vybe_check(ob_get_clean(), "99");
