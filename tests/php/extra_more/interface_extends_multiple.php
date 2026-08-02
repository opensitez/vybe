<?php
// vybe-test: php/extra_more/interface_extends_multiple
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

interface X2{public function x():int;}
interface Y2{public function y():int;}
interface Z2 extends X2,Y2{}
class Impl implements Z2{public function x():int{return 1;}public function y():int{return 2;}}
$o=new Impl; echo $o->x()+$o->y();

__vybe_check(ob_get_clean(), "3");
