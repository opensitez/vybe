<?php
// vybe-test: php/attributes/attribute_class_constructor_argument_read_via_reflection
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
class Version {
    public function __construct(public string $number) {}
}
#[Version('2.1.0')]
class App {}
echo (new ReflectionClass(App::class))->getAttributes(Version::class)[0]->newInstance()->number;

__vybe_check(ob_get_clean(), "2.1.0");
