<?php
// vybe-test: php/string_extra_builtins/parse_str_nested_brackets_runtime
// origin: languages/php/tests/php/test_string_extra_builtins.rs

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

parse_str('user[name]=john&user[age]=32&tags[]=a&tags[]=b', $out);
echo $out['user']['name'];
echo '|';
echo $out['user']['age'];
echo '|';
echo $out['tags'][0] . ',' . $out['tags'][1];

__vybe_check(ob_get_clean(), "john|32|a,b");
