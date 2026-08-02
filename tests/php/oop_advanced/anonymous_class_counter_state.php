<?php
// vybe-test: php/oop_advanced/anonymous_class_counter_state
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

function makeCounter(int $start = 0): object {
    return new class($start) {
        private int $value;
        public function __construct(int $start) { $this->value = $start; }
        public function inc(): void { $this->value++; }
        public function get(): int { return $this->value; }
    };
}
$c = makeCounter(10);
$c->inc();
$c->inc();
$c->inc();
echo $c->get(), "\n";

__vybe_check(ob_get_clean(), "13");
