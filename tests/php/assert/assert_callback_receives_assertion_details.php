<?php
// vybe-test: php/assert/assert_callback_receives_assertion_details
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

$seen = '';
$old = assert_options(ASSERT_CALLBACK, function($file, $line, $assertion, $desc = null) use (&$seen) {
    $seen = $line > 0 ? 'cb' : 'no';
    return true;
});
assert(false, 'via callback');
assert_options(ASSERT_CALLBACK, $old);
echo $seen;

__vybe_check(ob_get_clean(), "cb");
