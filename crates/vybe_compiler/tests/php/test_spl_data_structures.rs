use super::helpers::run_prints;

// ── SplStack (LIFO) ───────────────────────────────────────────

#[test]
fn splstack_push_and_pop() {
    assert_eq!(
        run_prints(
            r#"<?php
$s = new SplStack;
$s->push(1); $s->push(2); $s->push(3);
echo $s->pop() . ',' . $s->pop() . ',' . $s->pop();
"#
        ),
        vec!["3,2,1"]
    );
}
#[test]
fn splstack_top_peek() {
    assert_eq!(
        run_prints(
            r#"<?php
$s = new SplStack;
$s->push('a'); $s->push('b');
echo $s->top();
"#
        ),
        vec!["b"]
    );
}
#[test]
fn splstack_count() {
    assert_eq!(
        run_prints(
            r#"<?php
$s = new SplStack;
$s->push(1); $s->push(2);
echo count($s);
"#
        ),
        vec!["2"]
    );
}
#[test]
fn splstack_is_empty() {
    assert_eq!(
        run_prints(
            r#"<?php
$s = new SplStack;
echo $s->isEmpty() ? 'empty' : 'not';
$s->push(1);
echo $s->isEmpty() ? 'empty' : 'not';
"#
        ),
        vec!["empty", "not"]
    );
}

// ── SplQueue (FIFO) ───────────────────────────────────────────

#[test]
fn splqueue_enqueue_dequeue() {
    assert_eq!(
        run_prints(
            r#"<?php
$q = new SplQueue;
$q->enqueue('first'); $q->enqueue('second'); $q->enqueue('third');
echo $q->dequeue() . ',' . $q->dequeue();
"#
        ),
        vec!["first,second"]
    );
}
#[test]
fn splqueue_fifo_order() {
    assert_eq!(
        run_prints(
            r#"<?php
$q = new SplQueue;
foreach ([1,2,3,4,5] as $v) $q->enqueue($v);
$out = [];
while (!$q->isEmpty()) $out[] = $q->dequeue();
echo implode(',', $out);
"#
        ),
        vec!["1,2,3,4,5"]
    );
}
#[test]
fn splqueue_count() {
    assert_eq!(
        run_prints(
            r#"<?php
$q = new SplQueue;
$q->enqueue('a'); $q->enqueue('b');
echo count($q);
"#
        ),
        vec!["2"]
    );
}

// ── SplMinHeap / SplMaxHeap ───────────────────────────────────

#[test]
fn splminheap_extracts_minimum_first() {
    assert_eq!(
        run_prints(
            r#"<?php
$h = new SplMinHeap;
$h->insert(5); $h->insert(2); $h->insert(8); $h->insert(1);
$out = [];
while (!$h->isEmpty()) $out[] = $h->extract();
echo implode(',', $out);
"#
        ),
        vec!["1,2,5,8"]
    );
}
#[test]
fn splmaxheap_extracts_maximum_first() {
    assert_eq!(
        run_prints(
            r#"<?php
$h = new SplMaxHeap;
$h->insert(5); $h->insert(2); $h->insert(8); $h->insert(1);
$out = [];
while (!$h->isEmpty()) $out[] = $h->extract();
echo implode(',', $out);
"#
        ),
        vec!["8,5,2,1"]
    );
}
#[test]
fn splminheap_count() {
    assert_eq!(
        run_prints(
            r#"<?php
$h = new SplMinHeap;
$h->insert(3); $h->insert(1);
echo count($h);
"#
        ),
        vec!["2"]
    );
}

// ── SplFixedArray ─────────────────────────────────────────────

#[test]
fn splfixedarray_size_and_access() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = new SplFixedArray(5);
$a[0] = 10; $a[4] = 50;
echo $a[0] . ',' . $a[4] . ',' . count($a);
"#
        ),
        vec!["10,50,5"]
    );
}
#[test]
fn splfixedarray_from_array() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = SplFixedArray::fromArray([1,2,3,4,5]);
echo $a->getSize() . ':' . $a[2];
"#
        ),
        vec!["5:3"]
    );
}
#[test]
fn splfixedarray_to_array() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = SplFixedArray::fromArray([10,20,30]);
echo implode(',', $a->toArray());
"#
        ),
        vec!["10,20,30"]
    );
}
#[test]
fn splfixedarray_iteration() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = SplFixedArray::fromArray([5,3,1]);
$sum = 0;
foreach ($a as $v) $sum += $v;
echo $sum;
"#
        ),
        vec!["9"]
    );
}

// ── SplDoublyLinkedList ───────────────────────────────────────

#[test]
fn spldoublylinkedlist_push_pop() {
    assert_eq!(
        run_prints(
            r#"<?php
$l = new SplDoublyLinkedList;
$l->push('a'); $l->push('b'); $l->push('c');
echo $l->pop() . ',' . $l->shift();
"#
        ),
        vec!["c,a"]
    );
}
#[test]
fn spldoublylinkedlist_unshift() {
    assert_eq!(
        run_prints(
            r#"<?php
$l = new SplDoublyLinkedList;
$l->push('b'); $l->unshift('a');
echo $l->bottom() . ',' . $l->top();
"#
        ),
        vec!["a,b"]
    );
}
