<?php
// vybe-test: php/namespaces/global_namespace_block_accesses_both_namespaces
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

namespace N1 { class S { public function v(): int { return 1; } } }
namespace N2 { class S { public function v(): int { return 2; } } }
namespace {
    function sum(): int {
        return (new \N1\S())->v() + (new \N2\S())->v();
    }
}
echo sum();

__vybe_check(ob_get_clean(), "3");
