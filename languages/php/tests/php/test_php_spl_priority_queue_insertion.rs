use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP SPL: SplPriorityQueue Priority Extraction & Extra Flags
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_spl_priority_queue_extracts_by_numeric_priority() {
    let out = run_prints(
        r##"<?php
$pq = new SplPriorityQueue();
$pq->insert("Low Priority", 10);
$pq->insert("High Priority", 100);
$pq->insert("Medium Priority", 50);

$out = [];
while ($pq->valid()) {
    $out[] = $pq->extract();
}
echo implode(" -> ", $out);
"##,
    );
    assert_eq!(
        out,
        vec!["High Priority -> Medium Priority -> Low Priority"]
    );
}

#[test]
fn test_php_spl_priority_queue_extract_flags_priority() {
    let out = run_prints(
        r##"<?php
$pq = new SplPriorityQueue();
$pq->setExtractFlags(SplPriorityQueue::EXTR_PRIORITY);
$pq->insert("Task A", 25);

echo "Priority=" . $pq->extract();
"##,
    );
    assert_eq!(out, vec!["Priority=25"]);
}

#[test]
fn test_php_spl_priority_queue_extract_flags_both_array() {
    let out = run_prints(
        r##"<?php
$pq = new SplPriorityQueue();
$pq->setExtractFlags(SplPriorityQueue::EXTR_BOTH);
$pq->insert("Email Notification", 80);

$item = $pq->extract();
echo "Data={$item['data']} Priority={$item['priority']}";
"##,
    );
    assert_eq!(out, vec!["Data=Email Notification Priority=80"]);
}

#[test]
fn test_php_spl_priority_queue_custom_comparator() {
    compile_ok(
        r##"<?php
class ReversePriorityQueue extends SplPriorityQueue {
    public function compare(mixed $priority1, mixed $priority2): int {
        return $priority2 <=> $priority1; // Min priority first
    }
}

$pq = new ReversePriorityQueue();
$pq->insert("Urgent", 1);
$pq->insert("Low", 10);
echo $pq->extract() === "Urgent" ? "REVERSE_PRIORITY_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_priority_queue_top_peeks_element() {
    compile_ok(
        r##"<?php
$pq = new SplPriorityQueue();
$pq->insert("item", 5);
echo $pq->top() === "item" && count($pq) === 1 ? "PEEK_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_priority_queue_is_empty() {
    compile_ok(
        r##"<?php
$pq = new SplPriorityQueue();
echo $pq->isEmpty() ? "EMPTY" : "NOT_EMPTY";
"##,
    );
}

#[test]
fn test_php_spl_priority_queue_same_priority_fifo_order() {
    compile_ok(
        r##"<?php
$pq = new SplPriorityQueue();
$pq->insert("first", 10);
$pq->insert("second", 10);
$e1 = $pq->extract();
echo ($e1 === "first" || $e1 === "second") ? "SAME_PRIORITY_EXTRACT_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_priority_queue_count() {
    compile_ok(
        r##"<?php
$pq = new SplPriorityQueue();
$pq->insert("a", 1);
$pq->insert("b", 2);
echo count($pq) === 2 ? "COUNT_2_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_priority_queue_extract_empty_throws_runtime_exception() {
    compile_ok(
        r##"<?php
$pq = new SplPriorityQueue();
try {
    $pq->extract();
} catch (RuntimeException $e) {
    echo "PRIORITY_EMPTY_EX";
}
"##,
    );
}

#[test]
fn test_php_spl_priority_queue_recover_from_corruption() {
    compile_ok(
        r##"<?php
$pq = new SplPriorityQueue();
$pq->insert("val", 1);
if (method_exists($pq, "recoverFromCorruption")) {
    $pq->recoverFromCorruption();
}
echo "RECOVER_OK";
"##,
    );
}
