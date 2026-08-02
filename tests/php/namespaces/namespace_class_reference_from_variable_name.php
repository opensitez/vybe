<?php
// vybe-test: php/namespaces/namespace_class_reference_from_variable_name
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

namespace Domain {
    class Service {
        public function id(): string { return 'service'; }
    }
}
namespace App {
    $class = 'Domain\\Service';
    $svc = new $class();
    echo $svc->id();
}

__vybe_check(ob_get_clean(), "service");
