use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP SPL: SplStack LIFO Behavior, Inheritance & ArrayAccess
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_spl_stack_lifo_iteration_order() {
    let out = run_prints(
        r##"<?php
$stack = new SplStack();
$stack->push("bottom");
$stack->push("middle");
$stack->push("top");

$popped = [];
foreach ($stack as $item) {
    $popped[] = $item;
}
echo implode(" -> ", $popped);
"##,
    );
    assert_eq!(out, vec!["top -> middle -> bottom"]);
}

#[test]
fn test_php_spl_stack_inherits_spl_doubly_linked_list() {
    let out = run_prints(
        r##"<?php
$stack = new SplStack();
echo ($stack instanceof SplDoublyLinkedList) ? "IS_LINKED_LIST" : "FAIL";
"##,
    );
    assert_eq!(out, vec!["IS_LINKED_LIST"]);
}

#[test]
fn test_php_spl_stack_pop_removes_top_element() {
    let out = run_prints(
        r##"<?php
$stack = new SplStack();
$stack->push(10);
$stack->push(20);
$val = $stack->pop();
echo "Popped=$val Count=" . $stack->count();
"##,
    );
    assert_eq!(out, vec!["Popped=20 Count=1"]);
}

#[test]
fn test_php_spl_stack_default_iterator_mode_is_lifo_keep() {
    compile_ok(
        r##"<?php
$stack = new SplStack();
$stack->push("a");
$stack->push("b");
$mode = $stack->getIteratorMode();
echo ($mode & SplDoublyLinkedList::IT_MODE_LIFO) ? "MODE_LIFO" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_stack_offset_set_validates_bounds() {
    compile_ok(
        r##"<?php
$stack = new SplStack();
$stack->push("val1");
try {
    $stack[5] = "out_of_bounds";
} catch (OutOfRangeException $e) {
    echo "OutOfRangeException caught";
}
"##,
    );
}

#[test]
fn test_php_spl_stack_top_peek_without_popping() {
    compile_ok(
        r##"<?php
$s = new SplStack();
$s->push("element");
echo $s->top() === "element" && count($s) === 1 ? "PEEK_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_stack_array_access_indexing() {
    compile_ok(
        r##"<?php
$s = new SplStack();
$s->push("first_pushed");
$s->push("second_pushed");
// Index 0 in LIFO stack corresponds to top (second_pushed)
echo $s[0] === "second_pushed" ? "LIFO_INDEX0_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_stack_empty_check_with_is_empty() {
    compile_ok(
        r##"<?php
$s = new SplStack();
echo $s->isEmpty() ? "EMPTY" : "NOT_EMPTY";
"##,
    );
}

#[test]
fn test_php_spl_stack_push_multiple_types() {
    compile_ok(
        r##"<?php
$s = new SplStack();
$s->push(123);
$s->push(["key" => "value"]);
$s->push(new stdClass());
echo count($s) === 3 ? "PUSH_MIXED_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_stack_serialize_roundtrip() {
    compile_ok(
        r##"<?php
$s = new SplStack();
$s->push("data1");
$s->push("data2");
$serialized = serialize($s);
$restored = unserialize($serialized);
echo $restored->pop() === "data2" ? "RESTORED_LIFO_OK" : "FAIL";
"##,
    );
}
