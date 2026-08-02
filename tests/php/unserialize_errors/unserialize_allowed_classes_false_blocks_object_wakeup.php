<?php
// vybe-test: php/unserialize_errors/unserialize_allowed_classes_false_blocks_object_wakeup
// origin: languages/php/tests/php/test_unserialize_errors.rs

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

class Box { public int $n = 0; }
$blob = serialize(new Box());
$v = unserialize($blob, ['allowed_classes' => false]);
echo $v instanceof __PHP_Incomplete_Class ? 'blocked' : 'live';

__vybe_check(ob_get_clean(), "blocked");
