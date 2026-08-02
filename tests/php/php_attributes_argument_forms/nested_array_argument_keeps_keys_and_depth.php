<?php
// vybe-test: php/php_attributes_argument_forms/nested_array_argument_keeps_keys_and_depth
// origin: languages/php/tests/php/test_php_attributes_argument_forms.rs

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

#[Attribute]
class Cfg {
    public function __construct(public array $opts) {}
}
#[Cfg(['db' => ['host' => 'localhost', 'port' => 5432], 'debug' => true])]
class Settings {}
$o = (new ReflectionClass(Settings::class))->getAttributes(Cfg::class)[0]->newInstance()->opts;
echo $o['db']['host'] . ':' . $o['db']['port'] . ':' . ($o['debug'] ? 'on' : 'off');

__vybe_check(ob_get_clean(), "localhost:5432:on");
