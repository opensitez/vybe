<?php
// vybe-test: php/loops/foreach_over_generator_with_keys_and_break
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

function make_pairs(): Generator {
    yield 'a' => 1;
    yield 'b' => 2;
    yield 'c' => 3;
}
$out = '';
$count = 0;
foreach (make_pairs() as $k => $v) {
    $out .= $k . $v;
    $count++;
    if ($count === 2) {
        break;
    }
}
echo $out;

__vybe_check(ob_get_clean(), "a1b2");
