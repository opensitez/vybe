<?php
// vybe-test: php/php_autoloading/php_autoloading_class_exists_second_arg_toggle
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

$called = 0;
spl_autoload_register(function (string $class) use (&$called): void {
    if ($class === 'Autoload\\Maybe') {
        $called++;
        eval('class Autoload\\\\Maybe {}');
    }
});
echo class_exists('Autoload\\Maybe', false) ? 'pre' : 'pre-no';
echo $called . '|';
echo class_exists('Autoload\\Maybe', true) ? 'loaded' : 'noload';
echo $called . '|';
echo class_exists('Autoload\\Maybe', true) ? 'cached' : 'no-cache';
echo $called;

__vybe_check(ob_get_clean(), "pre-no1|loaded1|cached1");
