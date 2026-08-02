<?php
// vybe-test: php/php_class_alias_autoload_parameter/test_class_alias_instanceof_check
// origin: languages/php/tests/php/test_php_class_alias_autoload_parameter.rs

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

class Target {}
class_alias('Target', 'TargetAlias');
$a = new TargetAlias();
echo ($a instanceof Target) ? 'instance_of_target' : 'err', "\n";

__vybe_check(ob_get_clean(), "instance_of_target");
