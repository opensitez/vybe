<?php
// vybe-test: php/mbstring_extended/mb_detect_order_set_and_restore
// origin: languages/php/tests/php/test_mbstring_extended.rs

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

$old = mb_detect_order();
mb_detect_order(['ASCII', 'UTF-8']);
$new = mb_detect_order();
echo in_array('ASCII', $new, true) ? 'set' : 'unset';
mb_detect_order($old);
echo '|';
$restored = mb_detect_order();
echo is_array($restored) ? 'restored' : 'bad';

__vybe_check(ob_get_clean(), "set|restored");
