<?php
// vybe-test: php/php_reflection_class_method_property_attributes/test_php_reflection_attribute_reading
// origin: languages/php/tests/php/test_php_reflection_class_method_property_attributes.rs

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
class Entity {
    public function __construct(public string $table) {}
}

#[Entity("users_table")]
class User {}

$rc = new ReflectionClass(User::class);
$attrs = $rc->getAttributes(Entity::class);
$entity = $attrs[0]->newInstance();
echo $entity->table;

__vybe_check(ob_get_clean(), "users_table");
