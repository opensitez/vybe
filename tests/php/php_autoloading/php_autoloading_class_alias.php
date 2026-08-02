<?php
// vybe-test: php/php_autoloading/php_autoloading_class_alias
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

class AutoloadBase {}
class_alias(AutoloadBase::class, 'AutoloadAlias');
echo (new AutoloadAlias()) instanceof AutoloadBase ? 'yes' : 'no';
echo class_alias('AutoloadBase', 'AutoloadAlias', false) ? 'second' : 'first_fail';

__vybe_check(ob_get_clean(), "yesfirst_fail");
