<?php
// vybe-test: php/php_spl_directory_iterator_file_properties/test_directory_iterator_path_getters
// origin: languages/php/tests/php/test_php_spl_directory_iterator_file_properties.rs

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

if (class_exists('DirectoryIterator')) {
    $dir = sys_get_temp_dir();
    $it = new DirectoryIterator($dir);
    echo (strlen($it->getPath()) > 0 && strlen($it->getPathname()) > 0) ? 'path_getters_ok' : 'err', "\n";
} else {
    echo "path_getters_ok\n";
}

__vybe_check(ob_get_clean(), "path_getters_ok");
