<?php
// vybe-test: php/extra_more/trait_overrides_parent_method
// origin: languages/php/tests/php/test_extra_more.rs

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

class P{public function hello():string{return 'parent';}}
trait T{public function hello():string{return 'trait';}}
class C extends P{use T;}
echo (new C)->hello();

__vybe_check(ob_get_clean(), "trait");
