<?php
// vybe-test: php/php_autoloading/php_autoloading_no_autoload_when_flag_false
// origin: languages/php/tests/php/test_php_autoloading.rs

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
    if ($class === 'Autoload\\Missing') {
        eval('class Missing {}');
    }
});
echo class_exists('Autoload\\\\Missing', false) ? 'found' : 'not_found';
echo class_exists('Autoload\\\\Missing', true) ? 'loaded' : 'not_loaded';

__vybe_check(ob_get_clean(), "not_foundnot_loaded");
