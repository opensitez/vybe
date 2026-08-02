<?php
// vybe-test: php/reflection_attribute_instantiation/reflection_attribute_new_instance_named_args
// origin: languages/php/tests/php/test_reflection_attribute_instantiation.rs

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
class ConfigAttr {
    public function __construct(public bool $enabled, public int $level = 1) {}
}

#[ConfigAttr(level: 5, enabled: true)]
class Service {}

$rc = new ReflectionClass(Service::class);
$attr = $rc->getAttributes()[0];
$instance = $attr->newInstance();

echo $instance->enabled ? "yes" : "no";
echo "|" . $instance->level;

__vybe_check(ob_get_clean(), "yes|5");
