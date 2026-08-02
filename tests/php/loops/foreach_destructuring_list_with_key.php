<?php
// vybe-test: php/loops/foreach_destructuring_list_with_key
// origin: languages/php/tests/php/test_loops.rs

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

$items = [['k' => 'a', 'v' => 1], ['k' => 'b', 'v' => 2]];
$out = '';
foreach ($items as $idx => ['k' => $k, 'v' => $v]) {
    $out .= $idx . ':' . $k . $v . ';';
}
echo $out;

__vybe_check(ob_get_clean(), "0:a1;1:b2;");
