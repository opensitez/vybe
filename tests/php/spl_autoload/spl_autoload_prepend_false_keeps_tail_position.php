<?php
// vybe-test: php/spl_autoload/spl_autoload_prepend_false_keeps_tail_position
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

$hit = [];
spl_autoload_register(function (string $class) use (&$hit): void {
    if ($class === 'Pre\\Service') { $hit[] = 'first'; }
});
$loader = function (string $class) use (&$hit): void {
    if ($class === 'Pre\\Service') { $hit[] = 'second'; eval('namespace Pre; class Service {}'); }
};
spl_autoload_register($loader, true, false);
class_exists('Pre\\Service');
echo implode('|', $hit);

__vybe_check(ob_get_clean(), "first|second");
