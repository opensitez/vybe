<?php
// vybe-test: php/array_replace_recursive_deep/array_replace_recursive_removes_nesting_on_override
// origin: languages/php/tests/php/test_array_replace_recursive_deep.rs

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

$base = ['cfg' => ['mode' => ['safe' => true], 'level' => 1]];
$patch = ['cfg' => ['mode' => 'off']];
$res = array_replace_recursive($base, $patch);
echo is_array($res['cfg']['mode']) ? 'arr' : 'scalar';
echo "|" . $res['cfg']['mode'];

__vybe_check(ob_get_clean(), "scalar|off");
