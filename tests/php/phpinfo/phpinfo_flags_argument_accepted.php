<?php
// vybe-test: php/phpinfo/phpinfo_flags_argument_accepted
// origin: languages/php/tests/php/test_phpinfo.rs

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

phpinfo(INFO_GENERAL);
echo 'done';

__vybe_check(ob_get_clean(), "phpinfo()\nPHP Version => 8.0.0\nSystem => Darwin\nBuild Date => vybe\nServer API => cli\nPHP API => vybex\nPHP Extension Build => vybe\nZend Extension Build => n/a\nPHP Integer Size => 8\ndone");
