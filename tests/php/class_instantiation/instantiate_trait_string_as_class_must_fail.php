<?php
// vybe-test: php/class_instantiation/instantiate_trait_string_as_class_must_fail
// origin: languages/php/tests/php/test_class_instantiation.rs

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

trait TraitLike {}
$name = 'TraitLike';
try { new $name(); echo 'ok'; }
catch (Error $e) { echo 'trait-var'; }

__vybe_check(ob_get_clean(), "trait-var");
