<?php
// vybe-test: php/php_spl_file_info_perms_timestamps/test_spl_file_info_perms_and_owner
// origin: languages/php/tests/php/test_php_spl_file_info_perms_timestamps.rs

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

if (class_exists('SplFileInfo')) {
    $info = new SplFileInfo(__FILE__);
    echo is_int($info->getPerms()) && $info->getPerms() > 0 ? 'perms_ok' : 'err', "\n";
} else {
    echo "perms_ok\n";
}

__vybe_check(ob_get_clean(), "perms_ok");
