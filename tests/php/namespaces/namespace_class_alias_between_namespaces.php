<?php
// vybe-test: php/namespaces/namespace_class_alias_between_namespaces
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
    class Repo { public function type(): string { return 'repo'; } }
}
namespace App {
    use Domain\Repo as Repository;
    echo (new Repository())->type();
}

__vybe_check(ob_get_clean(), "repo");
