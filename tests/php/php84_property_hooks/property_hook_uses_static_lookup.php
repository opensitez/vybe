<?php
// vybe-test: php/php84_property_hooks/property_hook_uses_static_lookup
// origin: languages/php/tests/php/test_php84_property_hooks.rs

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
    private static array $data = [];
    public string $key {
        set(string $v) { $this->key = $v; self::$data[$v] = true; }
        get => $this->key;
    }
    public static function has(string $k): bool { return isset(self::$data[$k]); }
}
$r = new Registry();
$r->key = "mykey";
echo Registry::has("mykey") ? 'found' : 'not found';

__vybe_check(ob_get_clean(), "found");
