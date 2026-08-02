<?php
// vybe-test: php/fibers/generator_yield_key_preserved_in_foreach
// origin: languages/php/tests/php/test_fibers.rs

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

function gen_keys() {
    yield 0 => 'zero';
    yield 2 => 'two';
}
$out = [];
foreach (gen_keys() as $k => $v) {
    $out[] = $k . ':' . $v;
}
echo implode('|', $out);

__vybe_check(ob_get_clean(), "0:zero|2:two");
