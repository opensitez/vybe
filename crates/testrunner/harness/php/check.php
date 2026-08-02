<?php
// Vybe test harness — PHP.
//
// PHP does not pair like Go or JavaScript. `echo` is a statement, not a call,
// and it appends no newline: of 15,477 echos in the corpus only 1,248 carry a
// "\n", so consecutive echos land on ONE line and there is no i-th print to
// match to an i-th expected line.
//
// So instead of pairing, the whole program runs inside an output buffer and its
// entire output is compared once. That is exact rather than heuristic, and it
// needs no analysis of the program at all.
//
// The verdict is the EXIT CODE; the diagnostic is printed before throwing,
// because an uncaught error renders as `[object]` and says nothing.

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
