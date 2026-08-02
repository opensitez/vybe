<?php
// vybe-test: php/trait_conflict_resolution/class_using_trait_can_still_extend
// origin: languages/php/tests/php/test_trait_conflict_resolution.rs

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

trait Taggable { public function tag(): string { return "tagged"; } }
class Base { public function base(): string { return "base"; } }
class Child extends Base { use Taggable; }
$c = new Child();
echo $c->base() . ',' . $c->tag();

__vybe_check(ob_get_clean(), "base,tagged");
