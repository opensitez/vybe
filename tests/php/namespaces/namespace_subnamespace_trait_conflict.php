<?php
// vybe-test: php/namespaces/namespace_subnamespace_trait_conflict
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

namespace A {
    trait HasName { public function name(): string { return 'A'; } }
}
namespace B {
    trait HasName { public function name(): string { return 'B'; } }
}
namespace App {
    use A\HasName as AName;
    class C {
        use AName;
    }
    class D {
        use \B\HasName;
    }
}
echo (new \App\C())->name() . (new \App\D())->name();

__vybe_check(ob_get_clean(), "AB");
