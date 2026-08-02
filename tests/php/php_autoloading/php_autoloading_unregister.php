<?php
// vybe-test: php/php_autoloading/php_autoloading_unregister
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

$trace = [];
$loader = function (string $class) use (&$trace): void {
    $trace[] = $class;
    if ($class === 'Autoload\\Removable') {
        eval('class Autoload\\\\Removable {}');
    }
};
spl_autoload_register($loader);
echo class_exists('Autoload\\Removable', true) ? 'loaded' : 'not';
spl_autoload_unregister($loader);
echo class_exists('Autoload\\Never', true) ? 'bad' : 'missing';
echo implode('|', $trace);

__vybe_check(ob_get_clean(), "loadedmissing|Autoload\\\\Removable");
