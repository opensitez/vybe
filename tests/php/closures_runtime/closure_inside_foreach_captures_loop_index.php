<?php
// vybe-test: php/closures_runtime/closure_inside_foreach_captures_loop_index
// origin: languages/php/tests/php/test_closures_runtime.rs

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

$out = [];
foreach ([10, 20] as $i => $v) {
    $out[] = (function () use ($i, $v) { return $i . ':' . $v; })();
}
echo implode(',', $out);

__vybe_check(ob_get_clean(), "0:10,1:20");
