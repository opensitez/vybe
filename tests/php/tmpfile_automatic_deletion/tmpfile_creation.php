<?php
// vybe-test: php/tmpfile_automatic_deletion/tmpfile_creation
// origin: languages/php/tests/php/test_tmpfile_automatic_deletion.rs

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

$temp = tmpfile();
fwrite($temp, "test data");
rewind($temp);
echo fread($temp, 1024);
fclose($temp); // should delete the file

__vybe_check(ob_get_clean(), "test data");
