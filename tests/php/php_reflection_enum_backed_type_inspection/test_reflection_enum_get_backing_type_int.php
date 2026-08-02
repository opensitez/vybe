<?php
// vybe-test: php/php_reflection_enum_backed_type_inspection/test_reflection_enum_get_backing_type_int
// origin: languages/php/tests/php/test_php_reflection_enum_backed_type_inspection.rs

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

enum IntEnum: int {
    case One = 1;
}
$re = new ReflectionEnum(IntEnum::class);
$type = $re->getBackingType();
echo $type instanceof ReflectionNamedType ? $type->getName() : 'none', "\n";

__vybe_check(ob_get_clean(), "int");
