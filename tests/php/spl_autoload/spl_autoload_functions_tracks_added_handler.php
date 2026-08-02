<?php
// vybe-test: php/spl_autoload/spl_autoload_functions_tracks_added_handler
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

$before = function_exists('spl_autoload_functions') ? count((array) spl_autoload_functions()) : 0;
$loader = function (string $class): void {
    if ($class === 'Unused\\Class') {
        eval('namespace Unused; class Class {}');
    }
};
spl_autoload_register($loader);
$afterRegister = count((array) spl_autoload_functions());
spl_autoload_unregister($loader);
$afterUnregister = count((array) spl_autoload_functions());
echo ($afterRegister === $before + 1 ? 'plus1' : 'no');
echo '|';
echo ($afterUnregister === $before ? 'clean' : 'dirty');

__vybe_check(ob_get_clean(), "plus1|clean");
