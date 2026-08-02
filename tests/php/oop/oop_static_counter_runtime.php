<?php
// vybe-test: php/oop/oop_static_counter_runtime
// origin: languages/php/tests/php/test_oop.rs

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

class Stats {
    public static int $count = 0;
    public function tick(): int {
        return ++self::$count;
    }
}
$a = new Stats();
$b = new Stats();
echo $a->tick();
echo '|';
echo $b->tick();
echo '|';
echo Stats::$count;

__vybe_check(ob_get_clean(), "1|2|2");
