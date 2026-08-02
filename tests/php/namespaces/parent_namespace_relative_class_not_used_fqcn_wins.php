<?php
// vybe-test: php/namespaces/parent_namespace_relative_class_not_used_fqcn_wins
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

namespace Project\Core {
    class Engine { public function rev(): string { return 'v8'; } }
}
namespace Project\App {
    class Car {
        public function engine(): string {
            return (new \Project\Core\Engine())->rev();
        }
    }
}
echo (new \Project\App\Car())->engine();

__vybe_check(ob_get_clean(), "v8");
