<?php
// vybe-test: php/pcre_named_groups/preg_match_all_named_groups_set_order
// origin: languages/php/tests/php/test_pcre_named_groups.rs

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

preg_match_all('/(?P<k>\w+)=(?P<v>\d+)/', 'a=1 b=2 c=3', $m, PREG_SET_ORDER);
$pairs = array_map(fn($e) => $e['k'] . ':' . $e['v'], $m);
echo implode(',', $pairs);

__vybe_check(ob_get_clean(), "a:1,b:2,c:3");
