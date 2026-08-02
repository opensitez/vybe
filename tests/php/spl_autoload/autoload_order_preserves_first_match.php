<?php
// vybe-test: php/spl_autoload/autoload_order_preserves_first_match
// origin: languages/php/tests/php/test_spl_autoload.rs

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

$hits = [];
spl_autoload_register(function(string $class) use (&$hits): void {
    if ($class === 'Chain\\Svc') { $hits[] = 'first'; eval('namespace Chain; class Svc {}'); }
});
spl_autoload_register(function(string $class) use (&$hits): void {
    if ($class === 'Chain\\Svc') { $hits[] = 'second'; }
});
class_exists('Chain\\Svc');
echo implode('|', $hits);

__vybe_check(ob_get_clean(), "first");
