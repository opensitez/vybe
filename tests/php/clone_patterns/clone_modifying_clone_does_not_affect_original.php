<?php
// vybe-test: php/clone_patterns/clone_modifying_clone_does_not_affect_original
// origin: languages/php/tests/php/test_clone_patterns.rs

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

class Config { public string $host = 'localhost'; }
$orig = new Config();
$copy = clone $orig;
$copy->host = 'remote';
echo $orig->host . ',' . $copy->host;

__vybe_check(ob_get_clean(), "localhost,remote");
