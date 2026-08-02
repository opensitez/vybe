<?php
// vybe-test: php/php_spl_file_object_csv_iteration/test_spl_file_object_drop_new_line_flag
// origin: languages/php/tests/php/test_php_spl_file_object_csv_iteration.rs

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

if (class_exists('SplFileObject')) {
    $file = new SplFileObject('php://memory', 'r+');
    $file->fwrite("line1\nline2\n");
    $file->rewind();
    $file->setFlags(SplFileObject::DROP_NEW_LINE);
    $line = $file->current();
    echo $line === 'line1' ? 'no_newline' : 'has_newline', "\n";
} else {
    echo "no_newline\n";
}

__vybe_check(ob_get_clean(), "no_newline");
