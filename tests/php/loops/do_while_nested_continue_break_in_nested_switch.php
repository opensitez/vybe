<?php
// vybe-test: php/loops/do_while_nested_continue_break_in_nested_switch
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

$out = 0;
$i = 0;
do {
    switch ($i) {
        case 0:
            $i++;
            continue;
        case 1:
            $out += 2;
            break;
        default:
            $out += 3;
    }
    if ($i === 2) { break; }
    $i++;
} while ($i < 5);
echo $out . ':' . $i;

__vybe_check(ob_get_clean(), "5:2");
