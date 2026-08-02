<?php
// vybe-test: php/generators_patterns/generator_map_lazy
// origin: languages/php/tests/php/test_generators_patterns.rs

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

function mapGen(callable $fn, Generator $g): Generator { foreach ($g as $v) yield $fn($v); }
function genRange(int $a, int $b): Generator { for ($i=$a;$i<=$b;$i++) yield $i; }
$doubled = mapGen(fn($n) => $n*2, genRange(1, 5));
echo implode(',', iterator_to_array($doubled));

__vybe_check(ob_get_clean(), "2,4,6,8,10");
