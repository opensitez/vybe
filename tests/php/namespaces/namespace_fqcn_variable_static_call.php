<?php
// vybe-test: php/namespaces/namespace_fqcn_variable_static_call
// origin: languages/php/tests/php/test_namespaces.rs

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

namespace Services {
    class Factory {
        public static function make(int $n): int { return $n + 100; }
    }
}
namespace App {
    $class = '\\Services\\Factory';
    $method = 'make';
    echo $class::$method(9);
}

__vybe_check(ob_get_clean(), "109");
