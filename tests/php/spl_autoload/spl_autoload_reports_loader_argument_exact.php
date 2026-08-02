<?php
// vybe-test: php/spl_autoload/spl_autoload_reports_loader_argument_exact
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

$classes = [];
spl_autoload_register(function (string $class): void use (&$classes) {
    $classes[] = $class;
    if ($class === 'Exact\\Class') {
        eval('namespace Exact; class Class {}');
    }
});
class_exists('Exact\\Class');
echo $classes[0] === 'Exact\\Class' ? 'exact' : 'miss';

__vybe_check(ob_get_clean(), "exact");
