<?php
// vybe-test: php/namespaces/namespace_root_block_can_create_qualified_and_unqualified_classes
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
    class User { public function role(): string { return 'app'; } }
}
namespace {
    function make(): string {
        $user = new App\User();
        $plain = new \App\User();
        return $user->role() . '|' . $plain->role();
    }
    echo make();
}

__vybe_check(ob_get_clean(), "app|app");
