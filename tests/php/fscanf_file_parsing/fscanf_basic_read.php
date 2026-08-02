<?php
// vybe-test: php/fscanf_file_parsing/fscanf_basic_read
// origin: languages/php/tests/php/test_fscanf_file_parsing.rs

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

$fp = fopen("php://memory", "w+");
fwrite($fp, "101 John\n102 Jane");
rewind($fp);

$user1 = fscanf($fp, "%d %s");
$user2 = fscanf($fp, "%d %s");
echo $user1[0] . "-" . $user1[1] . "|" . $user2[0] . "-" . $user2[1];
fclose($fp);

__vybe_check(ob_get_clean(), "101-John|102-Jane");
