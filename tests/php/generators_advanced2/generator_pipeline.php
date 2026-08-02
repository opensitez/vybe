<?php
// vybe-test: php/generators_advanced2/generator_pipeline
// origin: languages/php/tests/php/test_generators_advanced2.rs

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

function doubled(Generator $g): Generator { foreach ($g as $v) yield $v * 2; }
function filtered(Generator $g, callable $fn): Generator { foreach ($g as $v) if ($fn($v)) yield $v; }
function gen(): Generator { for ($i=1;$i<=5;$i++) yield $i; }
$pipeline = filtered(doubled(gen()), fn($n) => $n > 4);
echo implode(',', iterator_to_array($pipeline));

__vybe_check(ob_get_clean(), "6,8,10");
