<?php
// vybe-test: php/loops/foreach_generator_continue_level_break_level
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

function range_iterable(int $to): Generator {
    for ($i = 0; $i <= $to; $i++) {
        yield $i;
    }
}
$out = '';
foreach (range_iterable(4) as $i) {
    if ($i === 1) { continue; }
    if ($i === 4) { break; }
    $out .= $i;
}
echo $out;

__vybe_check(ob_get_clean(), "023");
