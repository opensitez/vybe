<?php
// vybe-test: php/php81_features/intersection_type_accepted
// origin: languages/php/tests/php/test_php81_features.rs

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

interface Stringable2 { public function __toString(): string; }
interface Serializable2 { public function serialize(): string; }
class Item implements Stringable2, Serializable2 {
    public function __toString(): string { return 'item'; }
    public function serialize(): string { return 'serialized'; }
}
function process(Stringable2&Serializable2 $obj): string {
    return (string)$obj . ':' . $obj->serialize();
}
echo process(new Item);

__vybe_check(ob_get_clean(), "item:serialized");
