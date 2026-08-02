<?php
// vybe-test: php/anonymous_classes/anon_class_instanceof_interface
// origin: languages/php/tests/php/test_anonymous_classes.rs

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

interface Countable2 { public function count(): int; }
$obj = new class implements Countable2 { public function count(): int { return 5; } };
echo ($obj instanceof Countable2) ? 'yes' : 'no';

__vybe_check(ob_get_clean(), "yes");
