<?php
// vybe-test: php/generator_delegation_yield_from/generator_yield_from_generator
// origin: languages/php/tests/php/test_generator_delegation_yield_from.rs

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

function inner() { yield 'a'; yield 'b'; return 'c'; }
function outer() {
    $ret = yield from inner();
    yield $ret;
}
$out = [];
foreach (outer() as $v) $out[] = $v;
echo implode(',', $out);

__vybe_check(ob_get_clean(), "a,b,c");
