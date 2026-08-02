<?php
// vybe-test: php/php81_features/readonly_property_throws_on_reassign
// origin: languages/php/tests/php/test_php81_features.rs

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

class User { public function __construct(public readonly string $name) {} }
$u = new User('Alice');
try { $u->name = 'Bob'; } catch (Error $e) { echo 'readonly'; }

__vybe_check(ob_get_clean(), "readonly");
