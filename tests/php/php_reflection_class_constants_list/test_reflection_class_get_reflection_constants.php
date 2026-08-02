<?php
// vybe-test: php/php_reflection_class_constants_list/test_reflection_class_get_reflection_constants
// origin: languages/php/tests/php/test_php_reflection_class_constants_list.rs

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

class ConfigDemo {
    public const DEFAULT_HOST = '127.0.0.1';
    protected const PORT = 8080;
    private const SECRET = 'key123';
}
$rc = new ReflectionClass(ConfigDemo::class);
$consts = $rc->getReflectionConstants();
echo count($consts) . ':' . $consts[0]->getName(), "\n";

__vybe_check(ob_get_clean(), "3:DEFAULT_HOST");
