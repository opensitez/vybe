<?php
// vybe-test: php/spl_autoload/class_alias_name_collision_returns_false
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

class BaseThing {}
class_alias(BaseThing::class, 'AliasThing2');
$ok = class_alias(BaseThing::class, 'AliasThing2', false);
echo $ok ? 'created' : 'failed';

__vybe_check(ob_get_clean(), "failed");
