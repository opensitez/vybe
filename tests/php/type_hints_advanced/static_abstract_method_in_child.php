<?php
// vybe-test: php/type_hints_advanced/static_abstract_method_in_child
// origin: languages/php/tests/php/test_type_hints_advanced.rs

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
    abstract protected static function tableName(): string;
    public static function all(): string { return 'SELECT * FROM ' . static::tableName(); }
}
class Users extends Registry { protected static function tableName(): string { return 'users'; } }
echo Users::all();

__vybe_check(ob_get_clean(), "SELECT * FROM users");
