<?php
// vybe-test: php/covariant_return_types/static_return_type_returns_correct_class
// origin: languages/php/tests/php/test_covariant_return_types.rs

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
    private static array $items = [];
    public static function add(string $item): static {
        static::$items[] = $item;
        return new static();
    }
    public static function count(): int { return count(static::$items); }
}
Registry::add('a');
Registry::add('b');
echo Registry::count();

__vybe_check(ob_get_clean(), "2");
