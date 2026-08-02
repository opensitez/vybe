<?php
// vybe-test: php/spl_autoload/class_exists_false_skips_autoload_callback
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

$hit = 0;
$loader = function (string $class) use (&$hit): void {
    $hit++;
    if ($class === 'Skip\\Target') {
        eval('namespace Skip; class Target {}');
    }
};
spl_autoload_register($loader);
$result = class_exists('Skip\\Target', false) ? 'hit' : 'skip';
echo $result . '|' . $hit;
spl_autoload_unregister($loader);

__vybe_check(ob_get_clean(), "skip|0");
