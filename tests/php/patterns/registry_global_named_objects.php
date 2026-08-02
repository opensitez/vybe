<?php
// vybe-test: php/patterns/registry_global_named_objects
// origin: languages/php/tests/php/test_patterns.rs

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
    private static $instances = [];
    public static function set(string $key, $obj): void { self::$instances[$key] = $obj; }
    public static function get(string $key) { return self::$instances[$key] ?? null; }
    public static function has(string $key): bool { return isset(self::$instances[$key]); }
}
Registry::set('db', (object)['host' => 'localhost']);
echo Registry::has('db') ? 'found' : 'missing';
echo Registry::get('db')->host;
echo Registry::has('cache') ? 'found' : 'missing';

__vybe_check(ob_get_clean(), "foundlocalhostmissing");
