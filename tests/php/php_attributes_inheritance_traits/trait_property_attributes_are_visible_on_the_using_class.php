<?php
// vybe-test: php/php_attributes_inheritance_traits/trait_property_attributes_are_visible_on_the_using_class
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
class Column {
    public function __construct(public string $type) {}
}
trait HasId {
    #[Column('integer')]
    public $id;
}
class Thing {
    use HasId;
    #[Column('string')]
    public $name;
}
$out = [];
foreach ((new ReflectionClass(Thing::class))->getProperties() as $p) {
    foreach ($p->getAttributes(Column::class) as $a) {
        $out[] = $p->getName() . ':' . $a->newInstance()->type;
    }
}
echo implode('|', $out);

__vybe_check(ob_get_clean(), "name:string|id:integer");
