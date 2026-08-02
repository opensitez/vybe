<?php
// vybe-test: php/traits/trait_mutable_shared_state
// origin: languages/php/tests/php/test_traits.rs

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

trait Counter {
    public int $count = 0;
    public function inc(): int { return ++$this->count; }
}
class Worker {
    use Counter;
}
$a = new Worker();
$b = new Worker();
echo $a->inc() . $b->inc();

__vybe_check(ob_get_clean(), "11");
