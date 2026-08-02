<?php
// vybe-test: php/array_column_object_properties/array_column_objects_accessor_precedence
// origin: languages/php/tests/php/test_array_column_object_properties.rs

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

class WithBoth {
    private string $name;
    public function __construct(private int $id, string $name) {
        $this->name = $name;
    }
    public function __get($field) {
        return $field === 'name' ? 'magic-' . $this->name : null;
    }
    public function __isset($field): bool {
        return $field === 'name';
    }
}
$rows = [new WithBoth(1, 'X'), new WithBoth(2, 'Y')];
$vals = array_column($rows, 'name');
echo implode('|', $vals);

__vybe_check(ob_get_clean(), "magic-X|magic-Y");
