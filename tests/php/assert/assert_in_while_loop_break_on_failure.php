<?php
// vybe-test: php/assert/assert_in_while_loop_break_on_failure
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
$i = 0;
$out = '';
while ($i < 3) {
    try {
        assert($i !== 1);
        $out .= $i;
    } catch (AssertionError $e) {
        $out .= 'X';
        break;
    }
    $i++;
}
echo $out;

__vybe_check(ob_get_clean(), "0X");
