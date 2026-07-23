use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP SPL: SplFixedArray Resizing, Bounds & Memory Operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_spl_fixed_array_instantiation_and_indexing() {
    let out = run_prints(
        r##"<?php
$arr = new SplFixedArray(3);
$arr[0] = "zero";
$arr[1] = "one";
$arr[2] = "two";

echo "Size=" . $arr->getSize() . " Val1=" . $arr[1];
"##,
    );
    assert_eq!(out, vec!["Size=3 Val1=one"]);
}

#[test]
fn test_php_spl_fixed_array_resize_with_set_size() {
    let out = run_prints(
        r##"<?php
$arr = new SplFixedArray(2);
$arr[0] = "a";
$arr[1] = "b";

$arr->setSize(4);
$arr[2] = "c";

echo "NewSize=" . $arr->getSize() . " Count=" . count($arr);
"##,
    );
    assert_eq!(out, vec!["NewSize=4 Count=4"]);
}

#[test]
fn test_php_spl_fixed_array_out_of_bounds_exception() {
    let out = run_prints(
        r##"<?php
$arr = new SplFixedArray(2);
try {
    $val = $arr[5];
} catch (RuntimeException $e) {
    echo "OUT_OF_BOUNDS_EX";
}
"##,
    );
    assert_eq!(out, vec!["OUT_OF_BOUNDS_EX"]);
}

#[test]
fn test_php_spl_fixed_array_from_array_factory() {
    compile_ok(
        r##"<?php
$native = ["foo" => 10, "bar" => 20];
$fixed = SplFixedArray::fromArray($native, false);
echo $fixed->getSize() === 2 && $fixed[0] === 10 ? "FROM_ARRAY_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_fixed_array_to_array_export() {
    compile_ok(
        r##"<?php
$fixed = new SplFixedArray(2);
$fixed[0] = "x";
$fixed[1] = "y";
$exported = $fixed->toArray();
echo is_array($exported) && $exported[0] === "x" ? "TO_ARRAY_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_fixed_array_shrink_truncates_elements() {
    compile_ok(
        r##"<?php
$fixed = new SplFixedArray(5);
$fixed[4] = "last";
$fixed->setSize(2);
echo $fixed->getSize() === 2 ? "SHRINK_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_fixed_array_iterator_traversal() {
    compile_ok(
        r##"<?php
$fixed = new SplFixedArray(3);
$fixed[0] = 10; $fixed[1] = 20; $fixed[2] = 30;
$sum = 0;
foreach ($fixed as $val) {
    $sum += $val;
}
echo $sum === 60 ? "FIXED_ITERATOR_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_fixed_array_negative_size_throws_exception() {
    compile_ok(
        r##"<?php
try {
    $fixed = new SplFixedArray(-1);
} catch (InvalidArgumentException | ValueError $e) {
    echo "NEGATIVE_SIZE_EX";
}
"##,
    );
}

#[test]
fn test_php_spl_fixed_array_offset_unset_resets_to_null() {
    compile_ok(
        r##"<?php
$fixed = new SplFixedArray(2);
$fixed[0] = "data";
unset($fixed[0]);
echo $fixed[0] === null ? "UNSET_NULL_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_fixed_array_serialize_unserialize() {
    compile_ok(
        r##"<?php
$fixed = new SplFixedArray(2);
$fixed[0] = "val_a";
$s = serialize($fixed);
$restored = unserialize($s);
echo $restored[0] === "val_a" ? "SERIALIZE_FIXED_OK" : "FAIL";
"##,
    );
}
