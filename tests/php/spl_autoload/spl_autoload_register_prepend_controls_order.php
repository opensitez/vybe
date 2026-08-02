<?php
// vybe-test: php/spl_autoload/spl_autoload_register_prepend_controls_order
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

$log = [];
spl_autoload_register(function (string $class) use (&$log): void {
    if ($class === 'Order\\Widget') {
        $log[] = 'base';
        eval('namespace Order; class Widget {}');
    }
});
spl_autoload_register(function (string $class) use (&$log): void {
    if ($class === 'Order\\Widget') {
        $log[] = 'prepended';
    }
}, true, true);

if (class_exists('Order\\Widget')) {
    echo implode(',', $log) . '|loaded';
} else {
    echo implode(',', $log) . '|missing';
}

__vybe_check(ob_get_clean(), "prepended,base|loaded");
