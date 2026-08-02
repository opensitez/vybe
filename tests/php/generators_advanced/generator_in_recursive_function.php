<?php
// vybe-test: php/generators_advanced/generator_in_recursive_function
// origin: languages/php/tests/php/test_generators_advanced.rs

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

function permutations(array $items): Generator {
    if (count($items) <= 1) {
        yield $items;
        return;
    }
    foreach ($items as $k => $v) {
        $rest = $items;
        array_splice($rest, $k, 1);
        foreach (permutations($rest) as $perm) {
            yield array_merge([$v], $perm);
        }
    }
}
$count = 0;
foreach (permutations([1, 2, 3]) as $perm) {
    $count++;
}
echo $count; // 3! = 6

__vybe_check(ob_get_clean(), "6");
