<?php
// vybe-test: php/php_autoloading/php_autoloading_prepend_affects_loading_order
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

function autoload_order_default(string $class): void {
    if ($class === 'Autoload\\OrderProbe') {
        eval('class Autoload\\\\OrderProbe {}');
    }
}
function autoload_order_prepend(string $class): void {
    if ($class === 'Autoload\\OrderProbe') {
        eval('class Autoload\\\\OrderProbe {}');
    }
}
spl_autoload_register('autoload_order_default');
spl_autoload_register('autoload_order_prepend', true, true);
$functions = spl_autoload_functions();
if (is_array($functions) && count($functions) >= 2) {
    echo (is_array($functions[0]) ? $functions[0][0] : 'none') . '|';
    echo (is_array($functions[1]) ? $functions[1][0] : 'none');
} else {
    echo 'bad';
}

__vybe_check(ob_get_clean(), "autoload_order_prepend|autoload_order_default");
