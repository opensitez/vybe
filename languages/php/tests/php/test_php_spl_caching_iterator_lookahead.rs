use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP SPL: CachingIterator Lookahead, HasNext & ToString
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_spl_caching_iterator_has_next_lookahead() {
    let out = run_prints(
        r##"<?php
$arr = new ArrayIterator(["first", "second", "third"]);
$it = new CachingIterator($arr);

$results = [];
foreach ($it as $val) {
    $results[] = $val . ":" . ($it->hasNext() ? "HAS_MORE" : "LAST");
}
echo implode(" | ", $results);
"##,
    );
    assert_eq!(out, vec!["first:HAS_MORE | second:HAS_MORE | third:LAST"]);
}

#[test]
fn test_php_spl_caching_iterator_to_string_conversion() {
    let out = run_prints(
        r##"<?php
$arr = new ArrayIterator(["apple", "banana"]);
$it = new CachingIterator($arr, CachingIterator::TOSTRING_USE_CURRENT);

$it->rewind();
echo (string)$it;
"##,
    );
    assert_eq!(out, vec!["apple"]);
}

#[test]
fn test_php_spl_caching_iterator_full_cache_array_export() {
    let out = run_prints(
        r##"<?php
$arr = new ArrayIterator(["x" => 10, "y" => 20]);
$it = new CachingIterator($arr, CachingIterator::FULL_CACHE);

foreach ($it as $val) {}

$cache = $it->getCache();
echo "Cache count=" . count($cache) . " Y=" . $cache["y"];
"##,
    );
    assert_eq!(out, vec!["Cache count=2 Y=20"]);
}

#[test]
fn test_php_spl_caching_iterator_count_mode() {
    compile_ok(
        r##"<?php
$arr = new ArrayIterator([1, 2, 3, 4]);
$it = new CachingIterator($arr, CachingIterator::FULL_CACHE);
foreach ($it as $v) {}
echo count($it) === 4 ? "COUNT_CACHE_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_caching_iterator_get_inner_iterator() {
    compile_ok(
        r##"<?php
$inner = new ArrayIterator(["a", "b"]);
$it = new CachingIterator($inner);
echo $it->getInnerIterator() === $inner ? "INNER_ITERATOR_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_caching_iterator_offset_get_cache_key() {
    compile_ok(
        r##"<?php
$arr = new ArrayIterator(["key1" => "val1", "key2" => "val2"]);
$it = new CachingIterator($arr, CachingIterator::FULL_CACHE);
foreach ($it as $v) {}
echo $it["key1"] === "val1" ? "OFFSET_GET_CACHE_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_caching_iterator_catch_get_child_null() {
    compile_ok(
        r##"<?php
$arr = new ArrayIterator([1]);
$it = new CachingIterator($arr);
echo $it->getChildren() === null ? "NO_CHILDREN_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_caching_iterator_flags_setter_getter() {
    compile_ok(
        r##"<?php
$arr = new ArrayIterator(["test"]);
$it = new CachingIterator($arr);
$it->setFlags(CachingIterator::CALL_TOSTRING);
echo ($it->getFlags() & CachingIterator::CALL_TOSTRING) ? "FLAGS_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_caching_iterator_empty_inner_iterator() {
    compile_ok(
        r##"<?php
$arr = new ArrayIterator([]);
$it = new CachingIterator($arr);
echo !$it->hasNext() ? "EMPTY_HAS_NEXT_FALSE" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_caching_iterator_serialize_unserialize() {
    compile_ok(
        r##"<?php
$arr = new ArrayIterator(["x", "y"]);
$it = new CachingIterator($arr, CachingIterator::FULL_CACHE);
foreach ($it as $v) {}
$s = serialize($it);
$restored = unserialize($s);
echo count($restored) === 2 ? "SERIALIZE_CACHING_OK" : "FAIL";
"##,
    );
}
