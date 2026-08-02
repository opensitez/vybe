<?php
// vybe-test: php/generators_advanced/nested_generator_delegation
// origin: languages/php/tests/php/test_generators_advanced.rs

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

function level3() {
    yield "c";
    yield "d";
}
function level2() {
    yield "b";
    yield from level3();
    yield "e";
}
function level1() {
    yield "a";
    yield from level2();
    yield "f";
}
$result = [];
foreach (level1() as $v) {
    $result[] = $v;
}
echo implode("", $result);

__vybe_check(ob_get_clean(), "abcdef");
