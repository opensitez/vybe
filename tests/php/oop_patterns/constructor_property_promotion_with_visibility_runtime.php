<?php
// vybe-test: php/oop_patterns/constructor_property_promotion_with_visibility_runtime
// origin: languages/php/tests/php/test_oop_patterns.rs

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

class User {
    public function __construct(
        public string $name,
        protected int $age,
        private bool $active
    ) {}
    public function summary(): string {
        return $this->name . ':' . $this->age . ':' . ($this->active ? 'on' : 'off');
    }
}
$u = new User('alice', 31, true);
echo $u->name;
echo '|' . $u->summary();

__vybe_check(ob_get_clean(), "alice|alice:31:on");
