<?php
// vybe-test: php/oop_interfaces/interface_exists_for_declared_name_runtime
// origin: languages/php/tests/php/test_oop_interfaces.rs

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

interface ExistsDemo {
    public function ping(): string;
}
echo interface_exists('ExistsDemo') ? 'yes' : 'no';
echo '|';
echo interface_exists('DoesNotExist') ? 'yes' : 'no';

__vybe_check(ob_get_clean(), "yes|no");
