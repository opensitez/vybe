<?php
// vybe-test: php/cli_arguments/foreach_argv_collects_flags
// origin: languages/php/tests/php/test_cli_arguments.rs

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

$argv = ['tool.php', '-v', '-q', 'file.txt'];
$flags = '';
foreach ($argv as $i => $arg) {
    if ($i > 0 && str_starts_with($arg, '-')) {
        $flags .= $arg;
    }
}
echo $flags;

__vybe_check(ob_get_clean(), "-v-q");
