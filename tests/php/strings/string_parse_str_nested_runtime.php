<?php
// vybe-test: php/strings/string_parse_str_nested_runtime
// origin: languages/php/tests/php/test_strings.rs

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

$query = 'user[name]=alice&user[id]=7&tags[]=a&tags[]=b';
parse_str($query, $out);
echo $out['user']['name'];
echo '|';
echo $out['user']['id'];
echo '|';
echo implode(',', $out['tags']);

__vybe_check(ob_get_clean(), "alice|7|a,b");
