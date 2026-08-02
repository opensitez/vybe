<?php
// vybe-test: php/spl_autoload/get_declared_classes_reflects_autoloaded
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

spl_autoload_register(function (string $class): void {
    if ($class === 'AutoLoad\\Transient') {
        eval('namespace AutoLoad; class Transient {}');
    }
});
class_exists('AutoLoad\\Transient');
$declared = get_declared_classes();
echo in_array('AutoLoad\\Transient', $declared, true) ? 'yes' : 'no';

__vybe_check(ob_get_clean(), "yes");
