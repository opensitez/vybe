<?php
// vybe-test: php/loops/foreach_skips_missing_key_from_destructure
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

$pairs = [[ 'k' => 'a', 'v' => 1 ], [ 'k' => 'b' ]];
$out = [];
foreach ($pairs as $pair) {
    $out[] = $pair['k'] . ':' . ($pair['v'] ?? 0);
}
echo implode('|', $out);

__vybe_check(ob_get_clean(), "a:1|b:0");
