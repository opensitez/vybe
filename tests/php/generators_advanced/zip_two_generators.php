<?php
// vybe-test: php/generators_advanced/zip_two_generators
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

function zipGens(Generator $a, Generator $b) {
    while ($a->valid() && $b->valid()) {
        yield [$a->current(), $b->current()];
        $a->next();
        $b->next();
    }
}
function letters() {
    foreach (["a", "b", "c"] as $l) yield $l;
}
function numbers() {
    foreach ([1, 2, 3] as $n) yield $n;
}
$result = [];
foreach (zipGens(letters(), numbers()) as [$l, $n]) {
    $result[] = "$l$n";
}
echo implode(",", $result);

__vybe_check(ob_get_clean(), "a1,b2,c3");
