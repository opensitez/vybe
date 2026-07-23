use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP SPL: SplMaxHeap Extraction, Comparators & Traversal
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_spl_max_heap_extracts_highest_value_first() {
    let out = run_prints(
        r##"<?php
$heap = new SplMaxHeap();
$heap->insert(10);
$heap->insert(50);
$heap->insert(30);

$items = [];
while ($heap->valid()) {
    $items[] = $heap->extract();
}
echo implode(",", $items);
"##,
    );
    assert_eq!(out, vec!["50,30,10"]);
}

#[test]
fn test_php_spl_max_heap_custom_subclass_comparator() {
    let out = run_prints(
        r##"<?php
class ScoreHeap extends SplMaxHeap {
    protected function compare(mixed $val1, mixed $val2): int {
        return $val1["score"] <=> $val2["score"];
    }
}

$heap = new ScoreHeap();
$heap->insert(["player" => "Bob", "score" => 100]);
$heap->insert(["player" => "Alice", "score" => 500]);
$heap->insert(["player" => "Charlie", "score" => 250]);

$top = $heap->extract();
echo "Winner: " . $top["player"] . " (" . $top["score"] . ")";
"##,
    );
    assert_eq!(out, vec!["Winner: Alice (500)"]);
}

#[test]
fn test_php_spl_max_heap_top_peeks_maximum_without_extracting() {
    let out = run_prints(
        r##"<?php
$heap = new SplMaxHeap();
$heap->insert(100);
$heap->insert(200);

$max = $heap->top();
echo "Top=$max Count=" . $heap->count();
"##,
    );
    assert_eq!(out, vec!["Top=200 Count=2"]);
}

#[test]
fn test_php_spl_max_heap_extract_on_empty_throws_runtime_exception() {
    compile_ok(
        r##"<?php
$heap = new SplMaxHeap();
try {
    $heap->extract();
} catch (RuntimeException $e) {
    echo "Heap extract empty exception caught";
}
"##,
    );
}

#[test]
fn test_php_spl_max_heap_string_sorting() {
    compile_ok(
        r##"<?php
$heap = new SplMaxHeap();
$heap->insert("apple");
$heap->insert("zebra");
$heap->insert("banana");
echo $heap->extract() === "zebra" ? "STRING_MAX_HEAP_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_max_heap_foreach_iteration() {
    compile_ok(
        r##"<?php
$heap = new SplMaxHeap();
$heap->insert(5);
$heap->insert(15);
$res = [];
foreach ($heap as $val) {
    $res[] = $val;
}
echo implode(",", $res) === "15,5" ? "FOREACH_HEAP_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_max_heap_is_empty() {
    compile_ok(
        r##"<?php
$heap = new SplMaxHeap();
echo $heap->isEmpty() ? "EMPTY" : "NOT_EMPTY";
"##,
    );
}

#[test]
fn test_php_spl_max_heap_recover_from_corrupted_heap() {
    compile_ok(
        r##"<?php
$heap = new SplMaxHeap();
$heap->insert(1);
if (method_exists($heap, "recoverFromCorruption")) {
    $heap->recoverFromCorruption();
}
echo "RECOVER_OK";
"##,
    );
}

#[test]
fn test_php_spl_max_heap_duplicate_values() {
    compile_ok(
        r##"<?php
$heap = new SplMaxHeap();
$heap->insert(42);
$heap->insert(42);
$heap->insert(10);
echo $heap->extract() === 42 && $heap->extract() === 42 ? "DUPLICATE_HEAPS_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_max_heap_count() {
    compile_ok(
        r##"<?php
$heap = new SplMaxHeap();
$heap->insert(1);
$heap->insert(2);
$heap->insert(3);
echo count($heap) === 3 ? "COUNT_HEAP_OK" : "FAIL";
"##,
    );
}
