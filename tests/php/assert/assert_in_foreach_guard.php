<?php
// vybe-test: php/assert/assert_in_foreach_guard
// origin: languages/php/tests/php/test_assert.rs

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

assert_options(ASSERT_EXCEPTION, 1);
$log = [];
foreach ([1, 0, 2] as $n) {
    try {
        assert($n !== 0);
        $log[] = (string)$n;
    } catch (AssertionError $e) {
        $log[] = 'fail';
    }
}
echo implode(',', $log);

__vybe_check(ob_get_clean(), "1,fail,2");
