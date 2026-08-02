<?php
// vybe-test: php/sessions/session_unset_single_key_via_unset
// origin: languages/php/tests/php/test_sessions.rs

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

session_start();
$_SESSION['keep'] = 1;
$_SESSION['drop'] = 2;
unset($_SESSION['drop']);
echo isset($_SESSION['keep']) && !isset($_SESSION['drop']) ? 'ok' : 'no';

__vybe_check(ob_get_clean(), "ok");
