<?php
// vybe-test: php/intersection_types/class_satisfying_intersection_accepted_as_single_type
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

interface Printable2 { public function print2(): string; }
interface Saveable2 { public function save2(): string; }
class Doc implements Printable2, Saveable2 {
    public function print2(): string { return "print"; }
    public function save2(): string { return "save"; }
}
function useDoc(Printable2 $p): string { return $p->print2(); }
$doc = new Doc();
echo useDoc($doc);

__vybe_check(ob_get_clean(), "print");
