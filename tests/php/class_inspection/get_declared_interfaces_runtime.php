<?php
// vybe-test: php/class_inspection/get_declared_interfaces_runtime
// origin: languages/php/tests/php/test_class_inspection.rs

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

interface ListedInterface { public function run(): void; }
echo interface_exists('ListedInterface') ? 'yes' : 'no';
echo in_array('ListedInterface', get_declared_interfaces()) ? 'yes' : 'no';

__vybe_check(ob_get_clean(), "yesyes");
