<?php
// vybe-test: php/spl_autoload/spl_autoload_unregister_leaves_class_loaded
// origin: languages/php/tests/php/test_spl_autoload.rs

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

$loader = function (string $class): void {
    if ($class === 'Tmp\\Once') {
        eval('namespace Tmp; class Once {}');
    }
};
spl_autoload_register($loader);
class_exists('Tmp\\Once');
spl_autoload_unregister($loader);
echo class_exists('Tmp\\Once', false) ? 'loaded' : 'gone';

__vybe_check(ob_get_clean(), "loaded");
