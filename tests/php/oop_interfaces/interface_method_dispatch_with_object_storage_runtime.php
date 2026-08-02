<?php
// vybe-test: php/oop_interfaces/interface_method_dispatch_with_object_storage_runtime
// origin: languages/php/tests/php/test_oop_interfaces.rs

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

interface Strategy {
    public function execute(int $x, int $y): int;
}
class Add implements Strategy {
    public function execute(int $x, int $y): int { return $x + $y; }
}
class Mul implements Strategy {
    public function execute(int $x, int $y): int { return $x * $y; }
}
function run_all(array $items): int {
    $total = 0;
    foreach ($items as $item) { $total += $item->execute(3, 4); }
    return $total;
}
$items = [new Add(), new Mul()];
echo run_all($items);

__vybe_check(ob_get_clean(), "19");
