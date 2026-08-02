<?php
// vybe-test: php/typed_property_violations/readonly_class_all_properties_readonly
// origin: languages/php/tests/php/test_typed_property_violations.rs

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

readonly class Config { public function __construct(public int $port, public string $host) {} }
$c = new Config(8080, 'localhost');
echo $c->port . ':' . $c->host;

__vybe_check(ob_get_clean(), "8080:localhost");
