<?php
// vybe-test: php/php_attributes_argument_forms/omitted_default_is_absent_from_get_arguments_but_set_on_instance
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
class Opt {
    public function __construct(public int $a, public int $b = 9) {}
}
#[Opt(1)]
class One {}
$attr = (new ReflectionClass(One::class))->getAttributes(Opt::class)[0];
echo count($attr->getArguments()) . ':' . $attr->newInstance()->b;

__vybe_check(ob_get_clean(), "1:9");
