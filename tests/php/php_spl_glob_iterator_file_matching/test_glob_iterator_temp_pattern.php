<?php
// vybe-test: php/php_spl_glob_iterator_file_matching/test_glob_iterator_temp_pattern
// origin: languages/php/tests/php/test_php_spl_glob_iterator_file_matching.rs

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

if (class_exists('GlobIterator')) {
    $dir = sys_get_temp_dir();
    $it = new GlobIterator($dir . '/*');
    echo is_int($it->count()) && $it->count() >= 0 ? 'count_ok' : 'err', "\n";
} else {
    echo "count_ok\n";
}

__vybe_check(ob_get_clean(), "count_ok");
