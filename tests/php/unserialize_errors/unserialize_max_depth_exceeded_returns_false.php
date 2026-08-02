<?php
// vybe-test: php/unserialize_errors/unserialize_max_depth_exceeded_returns_false
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

$deep = ['l' => null];
$ref = &$deep['l'];
$ref = &$deep;
$blob = serialize($deep);
$v = @unserialize($blob, ['max_depth' => 2]);
echo $v === false ? 'depth' : 'parsed';

__vybe_check(ob_get_clean(), "depth");
