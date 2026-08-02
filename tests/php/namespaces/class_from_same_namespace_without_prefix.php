<?php
// vybe-test: php/namespaces/class_from_same_namespace_without_prefix
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

namespace App {
    class A { public function v(): int { return 1; } }
    class B {
        public function pull(): int {
            return (new A())->v();
        }
    }
}
echo (new \App\B())->pull();

__vybe_check(ob_get_clean(), "1");
