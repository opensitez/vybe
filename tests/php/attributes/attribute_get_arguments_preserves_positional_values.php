<?php
// vybe-test: php/attributes/attribute_get_arguments_preserves_positional_values
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
class Pair {
    public function __construct(public int $a, public int $b) {}
}
#[Pair(3, 7)]
class Box {}
$args = (new ReflectionClass(Box::class))->getAttributes(Pair::class)[0]->getArguments();
echo $args[0] . '+' . $args[1];

__vybe_check(ob_get_clean(), "3+7");
