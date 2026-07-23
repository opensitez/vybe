use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP SPL: SplDoublyLinkedList Core Operations & Traversal Modes
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_spl_doubly_linked_list_push_and_pop() {
    let out = run_prints(
        r##"<?php
$list = new SplDoublyLinkedList();
$list->push("first");
$list->push("second");
echo $list->pop() . " | remaining=" . $list->count();
"##,
    );
    assert_eq!(out, vec!["second | remaining=1"]);
}

#[test]
fn test_php_spl_doubly_linked_list_unshift_and_shift() {
    let out = run_prints(
        r##"<?php
$list = new SplDoublyLinkedList();
$list->unshift("a");
$list->unshift("b");
echo $list->shift() . " | " . $list->shift();
"##,
    );
    assert_eq!(out, vec!["b | a"]);
}

#[test]
fn test_php_spl_doubly_linked_list_iteration_mode_lifo() {
    let out = run_prints(
        r##"<?php
$list = new SplDoublyLinkedList();
$list->push(1);
$list->push(2);
$list->setIteratorMode(SplDoublyLinkedList::IT_MODE_LIFO | SplDoublyLinkedList::IT_MODE_KEEP);

$items = [];
foreach ($list as $val) {
    $items[] = $val;
}
echo implode(",", $items);
"##,
    );
    assert_eq!(out, vec!["2,1"]);
}

#[test]
fn test_php_spl_doubly_linked_list_iteration_mode_delete() {
    let out = run_prints(
        r##"<?php
$list = new SplDoublyLinkedList();
$list->push("x");
$list->push("y");
$list->setIteratorMode(SplDoublyLinkedList::IT_MODE_FIFO | SplDoublyLinkedList::IT_MODE_DELETE);

foreach ($list as $item) {}
echo "Count after delete iteration: " . $list->count();
"##,
    );
    assert_eq!(out, vec!["Count after delete iteration: 0"]);
}

#[test]
fn test_php_spl_doubly_linked_list_array_access_offset_set() {
    let out = run_prints(
        r##"<?php
$list = new SplDoublyLinkedList();
$list->push("alpha");
$list->push("beta");
$list[0] = "MODIFIED";
echo $list->bottom() . " -> " . $list->top();
"##,
    );
    assert_eq!(out, vec!["MODIFIED -> beta"]);
}

#[test]
fn test_php_spl_doubly_linked_list_empty_pop_throws_underflow() {
    compile_ok(
        r##"<?php
$list = new SplDoublyLinkedList();
try {
    $list->pop();
} catch (UnderflowException $e) {
    echo "UnderflowException caught";
}
"##,
    );
}

#[test]
fn test_php_spl_doubly_linked_list_add_at_index() {
    compile_ok(
        r##"<?php
$list = new SplDoublyLinkedList();
$list->push("A");
$list->push("C");
$list->add(1, "B");
echo $list[1] === "B" && count($list) === 3 ? "ADD_AT_INDEX_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_doubly_linked_list_offset_unset() {
    compile_ok(
        r##"<?php
$list = new SplDoublyLinkedList();
$list->push("item0");
$list->push("item1");
unset($list[0]);
echo count($list) === 1 && $list[0] === "item1" ? "OFFSET_UNSET_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_doubly_linked_list_serialize_unserialize() {
    compile_ok(
        r##"<?php
$list = new SplDoublyLinkedList();
$list->push(100);
$list->push(200);
$s = serialize($list);
$restored = unserialize($s);
echo count($restored) === 2 && $restored->top() === 200 ? "SERIALIZE_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_doubly_linked_list_bottom_and_top_getters() {
    compile_ok(
        r##"<?php
$list = new SplDoublyLinkedList();
$list->push("head");
$list->push("tail");
echo $list->bottom() === "head" && $list->top() === "tail" ? "BOTTOM_TOP_OK" : "FAIL";
"##,
    );
}
