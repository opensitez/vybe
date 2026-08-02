<?php
// vybe-test: php/extra_100/method_chaining_returns_this
// origin: languages/php/tests/php/test_extra_100.rs

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

class Str { private string $s=''; public function append(string $v):static{$this->s.=$v;return $this;} public function get():string{return $this->s;} }
echo (new Str)->append('a')->append('b')->append('c')->get();

__vybe_check(ob_get_clean(), "abc");
