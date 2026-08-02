<?php
// vybe-test: php/reflection_property_attributes/reflection_property_get_attributes
// origin: languages/php/tests/php/test_reflection_property_attributes.rs

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
class Column {
    public function __construct(public string $type) {}
}

class User {
    #[Column('integer')]
    public int $id;
    
    #[Column('string')]
    public string $name;
}

$rp1 = new ReflectionProperty(User::class, 'id');
$rp2 = new ReflectionProperty(User::class, 'name');

echo $rp1->getAttributes()[0]->getArguments()[0] . "|";
echo $rp2->getAttributes()[0]->getArguments()[0];

__vybe_check(ob_get_clean(), "integer|string");
