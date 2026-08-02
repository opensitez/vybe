<?php
// vybe-test: php/oop/oop_static_late_binding_runtime
// origin: languages/php/tests/php/test_oop.rs

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

class ServiceBase {
    public static string $scope = 'base';
    public static function scopeLabel(): string { return static::$scope; }
}
class ServiceChild extends ServiceBase {
    public static string $scope = 'child';
}
echo ServiceBase::scopeLabel();
echo '|';
echo ServiceChild::scopeLabel();

__vybe_check(ob_get_clean(), "base|child");
