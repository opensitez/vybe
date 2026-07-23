use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP SPL: SplQueue FIFO Operations — enqueue, dequeue, IteratorMode FIFO
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_spl_queue_fifo_enqueue_and_dequeue() {
    let out = run_prints(
        r##"<?php
$q = new SplQueue();
$q->enqueue("first_in");
$q->enqueue("second_in");

echo $q->dequeue() . " | " . $q->dequeue();
"##,
    );
    assert_eq!(out, vec!["first_in | second_in"]);
}

#[test]
fn test_php_spl_queue_fifo_iteration_order() {
    let out = run_prints(
        r##"<?php
$q = new SplQueue();
$q->enqueue(10);
$q->enqueue(20);
$q->enqueue(30);

$out = [];
foreach ($q as $v) {
    $out[] = $v;
}
echo implode(",", $out);
"##,
    );
    assert_eq!(out, vec!["10,20,30"]);
}

#[test]
fn test_php_spl_queue_dequeue_on_empty_throws_runtime_exception() {
    let out = run_prints(
        r##"<?php
$q = new SplQueue();
try {
    $q->dequeue();
} catch (RuntimeException $e) {
    echo "DEQUEUE_EMPTY_EX";
}
"##,
    );
    assert_eq!(out, vec!["DEQUEUE_EMPTY_EX"]);
}

#[test]
fn test_php_spl_queue_push_and_pop_aliases() {
    compile_ok(
        r##"<?php
$q = new SplQueue();
$q->push("a");
$q->push("b");
echo $q->pop() === "b" ? "POP_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_queue_bottom_peek_first_in() {
    compile_ok(
        r##"<?php
$q = new SplQueue();
$q->enqueue("job1");
$q->enqueue("job2");
echo $q->bottom() === "job1" ? "BOTTOM_JOB1_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_queue_count_tracking() {
    compile_ok(
        r##"<?php
$q = new SplQueue();
$q->enqueue(1);
$q->enqueue(2);
$q->dequeue();
echo count($q) === 1 ? "COUNT_1_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_queue_iterator_mode_delete_on_traversal() {
    compile_ok(
        r##"<?php
$q = new SplQueue();
$q->enqueue("msg1");
$q->enqueue("msg2");
$q->setIteratorMode(SplQueue::IT_MODE_DELETE);
foreach ($q as $msg) {}
echo count($q) === 0 ? "TRAVERSAL_DELETE_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_queue_array_access_offset_get() {
    compile_ok(
        r##"<?php
$q = new SplQueue();
$q->enqueue("x");
$q->enqueue("y");
echo $q[0] === "x" && $q[1] === "y" ? "FIFO_INDEX_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_queue_is_empty_check() {
    compile_ok(
        r##"<?php
$q = new SplQueue();
echo $q->isEmpty() ? "EMPTY_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_queue_serialize_unserialize() {
    compile_ok(
        r##"<?php
$q = new SplQueue();
$q->enqueue("task_alpha");
$s = serialize($q);
$restored = unserialize($s);
echo $restored->dequeue() === "task_alpha" ? "SERIALIZE_QUEUE_OK" : "FAIL";
"##,
    );
}
