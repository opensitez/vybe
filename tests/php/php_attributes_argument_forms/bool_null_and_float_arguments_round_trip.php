<?php
// vybe-test: php/php_attributes_argument_forms/bool_null_and_float_arguments_round_trip
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
class Types {
    public function __construct(public bool $flag, public ?string $none, public float $f) {}
}
#[Types(true, null, 1.5)]
class T {}
$i = (new ReflectionClass(T::class))->getAttributes(Types::class)[0]->newInstance();
echo var_export($i->flag, true) . ',' . var_export($i->none, true) . ',' . $i->f;

__vybe_check(ob_get_clean(), "true,NULL,1.5");
