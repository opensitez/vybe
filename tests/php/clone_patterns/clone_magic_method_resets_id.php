<?php
// vybe-test: php/clone_patterns/clone_magic_method_resets_id
// origin: languages/php/tests/php/test_clone_patterns.rs

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

class Entity {
    private static int $nextId = 1;
    public int $id;
    public function __construct() { $this->id = self::$nextId++; }
    public function __clone() { $this->id = self::$nextId++; }
}
$a = new Entity();
$b = clone $a;
echo $a->id . ',' . $b->id;

__vybe_check(ob_get_clean(), "1,2");
