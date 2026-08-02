<?php
// vybe-test: php/generator_nested_yield_from/generator_nested_yield_from
// origin: languages/php/tests/php/test_generator_nested_yield_from.rs

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

function gen3() { yield 3; }
function gen2() { yield 2; yield from gen3(); }
function gen1() { yield 1; yield from gen2(); yield 4; }

$out = [];
foreach (gen1() as $v) $out[] = $v;
echo implode(',', $out);

__vybe_check(ob_get_clean(), "1,2,3,4");
