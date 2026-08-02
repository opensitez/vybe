<?php
// vybe-test: php/generator_errors/foreach_stops_when_generator_throws_on_second_yield
// origin: languages/php/tests/php/test_generator_errors.rs

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

function gen(): Generator {
    yield 1;
    throw new RuntimeException('stop');
    yield 2;
}
$log = [];
try {
    foreach (gen() as $v) { $log[] = $v; }
} catch (RuntimeException $e) {
    $log[] = 'caught';
}
echo implode(',', $log);

__vybe_check(ob_get_clean(), "1,caught");
