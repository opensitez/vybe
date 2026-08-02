<?php
// vybe-test: php/intersection_types/intersection_requires_both_implementations
// origin: languages/php/tests/php/test_intersection_types.rs

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

interface Printable { public function print(): void; }
interface Saveable { public function save(): bool; }
class Document implements Printable, Saveable {
    public function print(): void { echo "printing"; }
    public function save(): bool { echo " saving"; return true; }
}
function process(Printable&Saveable $doc): void { $doc->print(); $doc->save(); }
process(new Document());

__vybe_check(ob_get_clean(), "printing saving");
