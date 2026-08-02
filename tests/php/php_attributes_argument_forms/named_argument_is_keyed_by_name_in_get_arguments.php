<?php
// vybe-test: php/php_attributes_argument_forms/named_argument_is_keyed_by_name_in_get_arguments
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
class Mix {
    public function __construct(public int $a, public string $b) {}
}
#[Mix(5, b: 'x')]
class M {}
$args = (new ReflectionClass(M::class))->getAttributes(Mix::class)[0]->getArguments();
echo $args[0] . ',' . $args['b'];

__vybe_check(ob_get_clean(), "5,x");
