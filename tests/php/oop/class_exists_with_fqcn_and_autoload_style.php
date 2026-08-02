<?php
// vybe-test: php/oop/class_exists_with_fqcn_and_autoload_style
// origin: languages/php/tests/php/test_oop.rs

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

class Demo {}
echo class_exists('Demo') ? 'yes' : 'no';
echo '|';
echo class_exists('\\Demo') ? 'yes' : 'no';

__vybe_check(ob_get_clean(), "yes|yes");
