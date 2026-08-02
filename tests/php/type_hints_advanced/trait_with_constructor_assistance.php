<?php
// vybe-test: php/type_hints_advanced/trait_with_constructor_assistance
// origin: languages/php/tests/php/test_type_hints_advanced.rs

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

trait AutoId {
    private static int $next = 0;
    private int $id;
    protected function initId(): void { $this->id = ++self::$next; }
    public function getId(): int { return $this->id; }
}
class Entity { use AutoId; public function __construct() { $this->initId(); } }
$a = new Entity; $b = new Entity; $c = new Entity;
echo $a->getId() . ',' . $b->getId() . ',' . $c->getId();

__vybe_check(ob_get_clean(), "1,2,3");
