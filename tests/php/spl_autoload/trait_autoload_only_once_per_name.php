<?php
// vybe-test: php/spl_autoload/trait_autoload_only_once_per_name
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

$seen = 0;
spl_autoload_register(function (string $class) use (&$seen): void {
    if ($class === 'AutoLoad\\TWidget') {
        $seen += 1;
        if ($seen === 1) {
            eval('namespace AutoLoad; trait TWidget {}');
        }
    }
});
trait_exists('AutoLoad\\TWidget');
trait_exists('AutoLoad\\TWidget');
echo $seen;

__vybe_check(ob_get_clean(), "1");
