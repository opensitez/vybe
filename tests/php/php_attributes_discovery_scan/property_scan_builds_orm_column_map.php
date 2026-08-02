<?php
// vybe-test: php/php_attributes_discovery_scan/property_scan_builds_orm_column_map
// origin: languages/php/tests/php/test_php_attributes_discovery_scan.rs

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
    public function __construct(public string $type, public bool $nullable = false) {}
}
#[Attribute]
class Id {}
class User {
    #[Id]
    #[Column('integer')]
    public int $id = 0;
    #[Column('string', nullable: true)]
    public ?string $email = null;
    public string $ignored = '';
}
$map = [];
foreach ((new ReflectionClass(User::class))->getProperties() as $p) {
    $cols = $p->getAttributes(Column::class);
    if (!$cols) continue;
    $c = $cols[0]->newInstance();
    $pk = $p->getAttributes(Id::class) ? '*' : '';
    $map[] = $pk . $p->getName() . ':' . $c->type . ($c->nullable ? '?' : '');
}
echo implode('|', $map);

__vybe_check(ob_get_clean(), "*id:integer|email:string?");
