<?php
// vybe-test: php/oop/oop_namespace_and_fqcn_runtime
// origin: languages/php/tests/php/test_oop.rs

namespace App\Module;

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

class Service {
    public function label(): string { return __CLASS__; }
}
echo (new Service())->label();
echo '|';
echo (new \App\Module\Service())->label();
echo '|';
echo class_exists('App\Module\Service') ? 'yes' : 'no';


__vybe_check(ob_get_clean(), "App\\Module\\Service|App\\Module\\Service|yes");
