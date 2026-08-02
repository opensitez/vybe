<?php
// vybe-test: php/namespaces/namespace_nested_resolution_with_current_namespace_prefix
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

namespace Infra {
    class Handler {
        public function run(): string { return 'run'; }
    }
}
namespace App\Runtime {
    function make(): string {
        $class = __NAMESPACE__ . '\\\\Handler';
        return (new $class())->run();
    }
    echo make();
}

__vybe_check(ob_get_clean(), "run");
