<?php
// vybe-test: php/php_autoloading/php_autoload_functions_list_contains_loader
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

$loader = function (string $class): void {
    if ($class === 'Autoload\\ListMe') { eval('class Autoload\\\\ListMe {}'); }
};
spl_autoload_register($loader);
$functions = spl_autoload_functions();
$found = false;
foreach ($functions as $f) {
    if (is_array($f) && $f[1] === '__invoke') {
        $found = true;
    }
}
echo $found ? 'found' : 'missing';
echo count($functions) >= 1 ? 'yes' : 'no';

__vybe_check(ob_get_clean(), "missingyes");
