<?php
// vybe-test: php/type_hints_advanced/dnf_union_of_intersections
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

interface X { public function x(): int; }
interface Y { public function y(): int; }
class Both implements X, Y { public function x(): int { return 1; } public function y(): int { return 2; } }
class OnlyX implements X { public function x(): int { return 10; } }
function sum((X&Y)|null $obj): int { return $obj === null ? 0 : $obj->x() + $obj->y(); }
echo sum(new Both) . ',' . sum(null);

__vybe_check(ob_get_clean(), "3,0");
