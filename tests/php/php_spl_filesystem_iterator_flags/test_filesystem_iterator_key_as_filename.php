<?php
// vybe-test: php/php_spl_filesystem_iterator_flags/test_filesystem_iterator_key_as_filename
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
    $it = new FilesystemIterator($dir, FilesystemIterator::KEY_AS_FILENAME | FilesystemIterator::SKIP_DOTS);
    $keysAreFilenames = true;
    foreach ($it as $key => $file) {
        if ($key !== $file->getFilename()) {
            $keysAreFilenames = false;
            break;
        }
    }
    echo $keysAreFilenames ? 'filename_keys_ok' : 'err', "\n";
} else {
    echo "filename_keys_ok\n";
}

__vybe_check(ob_get_clean(), "filename_keys_ok");
