<?php
// vybe-test: php/hash_functions/md5_file_equals_md5_of_contents
// origin: languages/php/tests/php/test_hash_functions.rs

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

$fp = fopen('php://memory', 'r+');
fwrite($fp, 'x');
rewind($fp);
$uri = stream_get_meta_data($fp)['uri'];
echo md5('x') === md5_file($uri) ? 'match' : 'diff';

__vybe_check(ob_get_clean(), "match");
