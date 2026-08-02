<?php
// vybe-test: php/typed_property_violations/readonly_property_second_assignment_fails
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

readonly class Token { public function __construct(public string $value) {} }
$t = new Token('abc');
try { $t->value = 'xyz'; echo 'mutated'; }
catch (Error $e) { echo 'readonly'; }

__vybe_check(ob_get_clean(), "readonly");
