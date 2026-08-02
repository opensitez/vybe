<?php
// vybe-test: php/php_oop_late_static_binding_self_static/test_php_late_static_binding_static_vs_self
// origin: languages/php/tests/php/test_php_oop_late_static_binding_self_static.rs

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

class BaseService {
    public static function getScopeSelf(): string { return self::getName(); }
    public static function getScopeStatic(): string { return static::getName(); }
    public static function getName(): string { return "Base"; }
}

class ChildService extends BaseService {
    public static function getName(): string { return "Child"; }
}

echo ChildService::getScopeSelf() . " vs " . ChildService::getScopeStatic();

__vybe_check(ob_get_clean(), "Base vs Child");
