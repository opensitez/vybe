<?php
// vybe-test: php/oop_advanced/static_property_per_subclass
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

class Registry {
    protected static array $items = [];
    public static function add(string $item): void {
        static::$items[] = $item;
    }
    public static function all(): array {
        return static::$items;
    }
}
class FruitRegistry extends Registry {
    protected static array $items = [];
}
class VegRegistry extends Registry {
    protected static array $items = [];
}
FruitRegistry::add("apple");
FruitRegistry::add("banana");
VegRegistry::add("carrot");
echo implode(",", FruitRegistry::all()), "\n";
echo implode(",", VegRegistry::all()), "\n";

__vybe_check(ob_get_clean(), "apple,banana\ncarrot");
