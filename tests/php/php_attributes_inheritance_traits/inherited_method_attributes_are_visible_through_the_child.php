<?php
// vybe-test: php/php_attributes_inheritance_traits/inherited_method_attributes_are_visible_through_the_child
// origin: languages/php/tests/php/test_php_attributes_inheritance_traits.rs

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
class Audit {
    public function __construct(public string $tag) {}
}
class Base {
    #[Audit('base-run')]
    public function run() {}
}
class Sub extends Base {}
$rm = new ReflectionMethod(Sub::class, 'run');
echo $rm->getAttributes(Audit::class)[0]->newInstance()->tag;

__vybe_check(ob_get_clean(), "base-run");
