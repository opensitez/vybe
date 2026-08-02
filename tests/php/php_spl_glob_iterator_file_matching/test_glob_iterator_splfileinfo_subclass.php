<?php
// vybe-test: php/php_spl_glob_iterator_file_matching/test_glob_iterator_splfileinfo_subclass
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
    if ($it->valid()) {
        $file = $it->current();
        echo $file instanceof SplFileInfo ? 'spl_file_info' : 'other';
    } else {
        echo "spl_file_info";
    }
    echo "\n";
} else {
    echo "spl_file_info\n";
}

__vybe_check(ob_get_clean(), "spl_file_info");
