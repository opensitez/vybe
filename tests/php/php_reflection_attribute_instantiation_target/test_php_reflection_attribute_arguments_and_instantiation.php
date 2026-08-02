<?php
// vybe-test: php/php_reflection_attribute_instantiation_target/test_php_reflection_attribute_arguments_and_instantiation
// origin: languages/php/tests/php/test_php_reflection_attribute_instantiation_target.rs

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

#[Attribute(Attribute::TARGET_CLASS)]
class Table {
    public function __construct(public string $name, public array $indexes = []) {}
}

#[Table("orders", indexes: ["idx_user_id"])]
class Order {}

$rc = new ReflectionClass(Order::class);
$attr = $rc->getAttributes(Table::class)[0];
$args = $attr->getArguments();
$instance = $attr->newInstance();

echo "Name={$instance->name} Index={$instance->indexes[0]} ArgCount=" . count($args);

__vybe_check(ob_get_clean(), "Name=orders Index=idx_user_id ArgCount=2");
