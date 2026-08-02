<?php
// vybe-test: php/attributes/attribute_get_arguments_array
// origin: languages/php/tests/php/test_attributes.rs

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
class Options {
    public function __construct(public array $list) {}
}
#[Options(['a' => 1, 'b' => 2])]
class ConfigHolder {}
$attr = (new ReflectionClass(ConfigHolder::class))->getAttributes(Options::class)[0];
$args = $attr->getArguments();
echo is_array($args) && isset($args[0]['a']) ? 'args_array_ok' : 'err';

__vybe_check(ob_get_clean(), "args_array_ok");
