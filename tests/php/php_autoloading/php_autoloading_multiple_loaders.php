<?php
// vybe-test: php/php_autoloading/php_autoloading_multiple_loaders
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
spl_autoload_register(function (string $class) use (&$trace): void {
    if ($class === 'Autoload\\A') {
        $trace[] = 'primary';
        eval('class Autoload\\A {}');
    }
}, true, false);
spl_autoload_register(function (string $class) use (&$trace): void {
    if ($class === 'Autoload\\A') {
        $trace[] = 'secondary';
        eval('class Autoload\\A2 {}');
    }
}, true, false);
$a = new Autoload\\A();
echo implode('|', $trace);
echo $a::class === 'Autoload\\\\A' ? 'ok' : 'bad';

__vybe_check(ob_get_clean(), "primaryok");
