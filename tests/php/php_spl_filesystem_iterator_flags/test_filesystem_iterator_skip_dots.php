<?php
// vybe-test: php/php_spl_filesystem_iterator_flags/test_filesystem_iterator_skip_dots
// origin: languages/php/tests/php/test_php_spl_filesystem_iterator_flags.rs

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

if (class_exists('FilesystemIterator')) {
    $dir = sys_get_temp_dir();
    $it = new FilesystemIterator($dir, FilesystemIterator::SKIP_DOTS);
    $foundDots = false;
    foreach ($it as $key => $file) {
        if ($file->getFilename() === '.' || $file->getFilename() === '..') {
            $foundDots = true;
        }
    }
    echo $foundDots ? 'has_dots' : 'no_dots', "\n";
} else {
    echo "no_dots\n";
}

__vybe_check(ob_get_clean(), "no_dots");
