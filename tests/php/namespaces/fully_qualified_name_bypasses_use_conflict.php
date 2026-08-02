<?php
// vybe-test: php/namespaces/fully_qualified_name_bypasses_use_conflict
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

namespace A { class Name { public function id(): string { return 'A'; } } }
namespace B { class Name { public function id(): string { return 'B'; } } }
namespace App {
    use A\Name;
    function pick(): string {
        $a = new Name();
        $b = new \B\Name();
        return $a->id() . $b->id();
    }
}
echo \App\pick();

__vybe_check(ob_get_clean(), "AB");
