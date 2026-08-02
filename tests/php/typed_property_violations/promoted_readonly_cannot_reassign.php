<?php
// vybe-test: php/typed_property_violations/promoted_readonly_cannot_reassign
// origin: languages/php/tests/php/test_typed_property_violations.rs

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

class User { public function __construct(public readonly int $id) {} }
$u = new User(1);
try { $u->id = 2; echo 'changed'; }
catch (Error $e) { echo 'blocked'; }

__vybe_check(ob_get_clean(), "blocked");
