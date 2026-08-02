<?php
// vybe-test: php/abstract_final_patterns/abstract_class_static_method_callable
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

abstract class Registry {
    private static array $items = [];
    public static function add(string $item): void { self::$items[] = $item; }
    public static function all(): array { return self::$items; }
}
class MyRegistry extends Registry {}
MyRegistry::add("a");
MyRegistry::add("b");
echo implode(',', MyRegistry::all()), "\n";

__vybe_check(ob_get_clean(), "a,b");
