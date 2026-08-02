<?php
// vybe-test: php/generator_errors/generator_delegates_return_from_inner_via_yield_from
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

function inner(): Generator { yield 1; return 9; }
function outer(): Generator { return yield from inner(); }
$g = outer();
$g->next();
$g->next();
try { echo $g->getReturn(); } catch (Exception $e) { echo 'no'; }

__vybe_check(ob_get_clean(), "9");
