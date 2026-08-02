<?php
// vybe-test: php/generator_memory_cleanup/generator_memory_cleanup
// origin: languages/php/tests/php/test_generator_memory_cleanup.rs

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

class ResourceObj {
    public static $count = 0;
    public function __construct() { self::$count++; }
    public function __destruct() { self::$count--; }
}

function gen() {
    $obj = new ResourceObj();
    yield 1;
    yield 2;
}
$g = gen();
$g->current();
echo ResourceObj::$count . "|";
$g = null;
echo ResourceObj::$count;

__vybe_check(ob_get_clean(), "1|0");
