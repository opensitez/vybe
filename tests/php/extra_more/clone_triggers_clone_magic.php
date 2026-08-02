<?php
// vybe-test: php/extra_more/clone_triggers_clone_magic
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

class C{public int $v=0;public function __clone(){$this->v++;}}
$a=new C; $b=clone $a;
echo $b->v;

__vybe_check(ob_get_clean(), "1");
