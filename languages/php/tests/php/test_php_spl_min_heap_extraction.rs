use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP SPL: SplMinHeap Extraction & Ascending Sort Behavior
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_spl_min_heap_extracts_lowest_value_first() {
    let out = run_prints(
        r##"<?php
$heap = new SplMinHeap();
$heap->insert(50);
$heap->insert(10);
$heap->insert(30);

$items = [];
while ($heap->valid()) {
    $items[] = $heap->extract();
}
echo implode(",", $items);
"##,
    );
    assert_eq!(out, vec!["10,30,50"]);
}

#[test]
fn test_php_spl_min_heap_top_peeks_minimum() {
    let out = run_prints(
        r##"<?php
$heap = new SplMinHeap();
$heap->insert(99);
$heap->insert(5);
$heap->insert(42);

echo "Min: " . $heap->top();
"##,
    );
    assert_eq!(out, vec!["Min: 5"]);
}

#[test]
fn test_php_spl_min_heap_custom_date_comparator() {
    let out = run_prints(
        r##"<?php
class DateMinHeap extends SplMinHeap {
    protected function compare(mixed $val1, mixed $val2): int {
        return strtotime($val1) <=> strtotime($val2);
    }
}

$heap = new DateMinHeap();
$heap->insert("2024-12-31");
$heap->insert("2024-01-01");
$heap->insert("2024-06-15");

echo "Earliest date: " . $heap->extract();
"##,
    );
    assert_eq!(out, vec!["Earliest date: 2024-01-01"]);
}

#[test]
fn test_php_spl_min_heap_is_empty_state() {
    compile_ok(
        r##"<?php
$heap = new SplMinHeap();
echo $heap->isEmpty() ? "IS_EMPTY_TRUE" : "FAIL";
$heap->insert(1);
echo !$heap->isEmpty() ? " IS_EMPTY_FALSE" : " FAIL";
"##,
    );
}

#[test]
fn test_php_spl_min_heap_count_after_extractions() {
    compile_ok(
        r##"<?php
$heap = new SplMinHeap();
$heap->insert(100);
$heap->insert(200);
$heap->extract();
echo count($heap) === 1 ? "COUNT_1_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_min_heap_duplicate_numbers() {
    compile_ok(
        r##"<?php
$heap = new SplMinHeap();
$heap->insert(7);
$heap->insert(7);
$heap->insert(2);
echo $heap->extract() === 2 && $heap->extract() === 7 ? "MIN_DUPLICATES_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_min_heap_string_alphabetical_ascending() {
    compile_ok(
        r##"<?php
$heap = new SplMinHeap();
$heap->insert("charlie");
$heap->insert("alice");
$heap->insert("bob");
echo $heap->extract() === "alice" ? "ALICE_MIN_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_min_heap_rewind_restarts_iteration() {
    compile_ok(
        r##"<?php
$heap = new SplMinHeap();
$heap->insert(10);
$heap->insert(20);
$heap->rewind();
echo $heap->current() === 10 ? "REWIND_MIN_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_min_heap_extract_empty_throws_runtime_exception() {
    compile_ok(
        r##"<?php
$heap = new SplMinHeap();
try {
    $heap->extract();
} catch (RuntimeException $e) {
    echo "MIN_HEAP_EMPTY_EX";
}
"##,
    );
}

#[test]
fn test_php_spl_min_heap_float_comparisons() {
    compile_ok(
        r##"<?php
$heap = new SplMinHeap();
$heap->insert(3.14);
$heap->insert(1.41);
$heap->insert(2.71);
echo $heap->extract() === 1.41 ? "FLOAT_MIN_OK" : "FAIL";
"##,
    );
}
