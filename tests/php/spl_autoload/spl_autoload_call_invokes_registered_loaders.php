<?php
// vybe-test: php/spl_autoload/spl_autoload_call_invokes_registered_loaders
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

$log = [];
$loader = function (string $class) use (&$log): void {
    if ($class === 'Manual\\Probe') {
        $log[] = 'called';
        eval('namespace Manual; class Probe {}');
    }
};
spl_autoload_register($loader);
spl_autoload_call('Manual\\Probe');
echo (class_exists('Manual\\Probe', false) ? 'exists' : 'missing') . '|' . implode(',', $log);

__vybe_check(ob_get_clean(), "exists|called");
