<?php
// vybe-test: php/oop_advanced/static_accumulator_across_instances
// origin: languages/php/tests/php/test_oop_advanced.rs

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

class Counter {
    private static int $total = 0;
    private int $id;
    public function __construct() {
        self::$total++;
        $this->id = self::$total;
    }
    public function getId(): int { return $this->id; }
    public static function getTotal(): int { return self::$total; }
}
$a = new Counter();
$b = new Counter();
$c = new Counter();
echo $a->getId(), "\n";
echo $b->getId(), "\n";
echo $c->getId(), "\n";
echo Counter::getTotal(), "\n";

__vybe_check(ob_get_clean(), "1\n2\n3\n3");
