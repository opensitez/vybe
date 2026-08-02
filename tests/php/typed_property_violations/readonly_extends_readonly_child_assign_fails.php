<?php
// vybe-test: php/typed_property_violations/readonly_extends_readonly_child_assign_fails
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

readonly class Base { public function __construct(public int $v) {} }
readonly class Child extends Base {}
$c = new Child(5);
try { $c->v = 6; echo 'ok'; }
catch (Error $e) { echo 'fail'; }

__vybe_check(ob_get_clean(), "fail");
