<?php
// vybe-test: php/spl_autoload/autoload_unregister_removes_loader
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

$count = 0;
$loader = function(string $class) use (&$count): void {
    if ($class === 'Temp\\Svc') {
        $count++;
        eval('namespace Temp; class Svc {}');
    }
};
spl_autoload_register($loader);
spl_autoload_unregister($loader);
class_exists('Temp\\Svc');
echo $count;

__vybe_check(ob_get_clean(), "0");
