<?php
// vybe-test: php/php_attributes_argument_forms/enum_case_argument_survives_new_instance
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

enum ColumnType: string {
    case Str = 'string';
    case Int = 'integer';
}
#[Attribute]
class Column {
    public function __construct(public ColumnType $type) {}
}
class Row {
    #[Column(ColumnType::Str)]
    public $name;
}
$a = (new ReflectionProperty(Row::class, 'name'))->getAttributes(Column::class)[0];
echo $a->newInstance()->type->value;

__vybe_check(ob_get_clean(), "string");
