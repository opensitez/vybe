<?php
// vybe-test: php/fscanf_file_parsing/fscanf_pass_by_reference
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
fwrite($fp, "Color: Red\n");
rewind($fp);

$count = fscanf($fp, "Color: %s", $color);
echo $count . "|" . $color;
fclose($fp);

__vybe_check(ob_get_clean(), "1|Red");
