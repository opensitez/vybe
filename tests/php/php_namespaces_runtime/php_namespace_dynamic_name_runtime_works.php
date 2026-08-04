<?php
// vybe-test: php/php_namespaces_runtime/php_namespace_dynamic_name_runtime_works
// origin: languages/php/tests/php/test_php_namespaces_runtime.rs

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

namespace Runtime;
class Service { public function endpoint(): string { return 'ok'; } }
$name = 'Runtime\\Service';
$class = new $name();
echo $class->endpoint();

__vybe_check(ob_get_clean(), "ok");
