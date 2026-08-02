<?php
// vybe-test: php/regex_patterns/named_groups_all
// origin: languages/php/tests/php/test_regex_patterns.rs

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

$log = "ERROR: file not found\nWARN: low memory\nERROR: timeout";
preg_match_all('/(?P<level>ERROR|WARN): (?P<msg>.+)/', $log, $matches);
echo implode(",", $matches['level']);
echo implode(",", $matches['msg']);

__vybe_check(ob_get_clean(), "ERROR,WARN,ERRORfile not found,low memory,timeout");
