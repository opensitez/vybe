<?php
// vybe-test: php/php_attributes_container_wiring/get_target_reports_where_the_attribute_was_applied
// origin: languages/php/tests/php/test_php_attributes_container_wiring.rs

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

#[Attribute(Attribute::TARGET_CLASS | Attribute::TARGET_METHOD)]
class Tag {}
#[Tag]
class Klass {
    #[Tag]
    public function m() {}
}
$ct = (new ReflectionClass(Klass::class))->getAttributes(Tag::class)[0]->getTarget();
$mt = (new ReflectionMethod(Klass::class, 'm'))->getAttributes(Tag::class)[0]->getTarget();
echo $ct . ',' . $mt . ',' . Attribute::TARGET_CLASS . ',' . Attribute::TARGET_METHOD;

__vybe_check(ob_get_clean(), "1,4,1,4");
