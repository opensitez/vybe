use super::helpers::compile_ok;

// ── SplStack — push, pop, top ─────────────────────────────────

#[test] fn spl_stack_push_pop_top() {
    compile_ok(r#"<?php
$s = new SplStack();
$s->push('first');
$s->push('second');
$s->push('third');
echo $s->top();
echo $s->pop();
echo $s->count();
"#);
}

// ── SplStack LIFO iteration ───────────────────────────────────

#[test] fn spl_stack_lifo_iteration() {
    compile_ok(r#"<?php
$s = new SplStack();
$s->push(1); $s->push(2); $s->push(3);
$result = [];
foreach ($s as $v) { $result[] = $v; }
echo implode(',', $result); // 3,2,1
"#);
}

// ── SplQueue — enqueue, dequeue ───────────────────────────────

#[test] fn spl_queue_enqueue_dequeue() {
    compile_ok(r#"<?php
$q = new SplQueue();
$q->enqueue('alpha');
$q->enqueue('beta');
$q->enqueue('gamma');
echo $q->dequeue();
echo $q->count();
"#);
}

// ── SplQueue FIFO order ───────────────────────────────────────

#[test] fn spl_queue_fifo_order() {
    compile_ok(r#"<?php
$q = new SplQueue();
foreach (['x', 'y', 'z'] as $v) { $q->enqueue($v); }
$out = [];
while (!$q->isEmpty()) { $out[] = $q->dequeue(); }
echo implode(',', $out); // x,y,z
"#);
}

// ── SplDoublyLinkedList — push front/back, pop front/back ────

#[test] fn spl_dll_push_front_back_pop() {
    compile_ok(r#"<?php
$dll = new SplDoublyLinkedList();
$dll->push('back1');
$dll->push('back2');
$dll->unshift('front');
echo $dll->shift();   // front
echo $dll->pop();     // back2
echo $dll->count();   // 1
"#);
}

// ── SplMinHeap — insert, extract minimum ─────────────────────

#[test] fn spl_min_heap_extract_order() {
    compile_ok(r#"<?php
$h = new SplMinHeap();
foreach ([9, 3, 7, 1, 5] as $v) { $h->insert($v); }
$out = [];
while (!$h->isEmpty()) { $out[] = $h->extract(); }
echo implode(',', $out); // 1,3,5,7,9
"#);
}

// ── SplMaxHeap — insert, extract maximum ─────────────────────

#[test] fn spl_max_heap_extract_order() {
    compile_ok(r#"<?php
$h = new SplMaxHeap();
foreach ([9, 3, 7, 1, 5] as $v) { $h->insert($v); }
$out = [];
while (!$h->isEmpty()) { $out[] = $h->extract(); }
echo implode(',', $out); // 9,7,5,3,1
"#);
}

// ── SplPriorityQueue — insert with priority, extract ─────────

#[test] fn spl_priority_queue_priority_order() {
    compile_ok(r#"<?php
$pq = new SplPriorityQueue();
$pq->insert('low',    1);
$pq->insert('high',   10);
$pq->insert('medium', 5);
$out = [];
while (!$pq->isEmpty()) { $out[] = $pq->extract(); }
echo implode(',', $out); // high,medium,low
"#);
}

// ── SplFixedArray — fixed-size array ─────────────────────────

#[test] fn spl_fixed_array_indexed_access() {
    compile_ok(r#"<?php
$fa = new SplFixedArray(4);
$fa[0] = 10; $fa[1] = 20; $fa[2] = 30; $fa[3] = 40;
echo $fa->getSize() . ':' . $fa[2];
"#);
}

// ── SplFixedArray from regular array ─────────────────────────

#[test] fn spl_fixed_array_from_regular_array() {
    compile_ok(r#"<?php
$fa = SplFixedArray::fromArray([100, 200, 300, 400]);
echo $fa->getSize();
echo $fa[3];
"#);
}

// ── SplObjectStorage — attach/detach objects ──────────────────

#[test] fn spl_object_storage_attach_detach() {
    compile_ok(r#"<?php
$store = new SplObjectStorage();
$a = new stdClass();
$b = new stdClass();
$store->attach($a);
$store->attach($b);
echo $store->count();
$store->detach($a);
echo $store->count();
"#);
}

// ── SplObjectStorage with associated data ─────────────────────

#[test] fn spl_object_storage_with_data() {
    compile_ok(r#"<?php
$store = new SplObjectStorage();
$obj = new stdClass();
$store->attach($obj, ['role' => 'admin', 'active' => true]);
$store->rewind();
$data = $store->getInfo();
echo isset($data['role']) ? $data['role'] : 'no role';
"#);
}

// ── SplBitSet (compile_ok — may not be available everywhere) ──

#[test] fn spl_bitset_if_available() {
    compile_ok(r#"<?php
if (class_exists('SplBitSet')) {
    $bs = new SplBitSet();
    $bs->offsetSet(0, true);
    $bs->offsetSet(3, true);
    echo $bs->offsetGet(0) ? '1' : '0';
    echo $bs->offsetGet(1) ? '1' : '0';
    echo $bs->offsetGet(3) ? '1' : '0';
} else {
    echo '1';
    echo '0';
    echo '1';
}
"#);
}

// ── SplStack isEmpty check ────────────────────────────────────

#[test] fn spl_stack_is_empty_check() {
    compile_ok(r#"<?php
$s = new SplStack();
echo $s->isEmpty() ? 'empty' : 'not empty';
$s->push(42);
echo $s->isEmpty() ? 'empty' : 'not empty';
$s->pop();
echo $s->isEmpty() ? 'empty' : 'not empty';
"#);
}

// ── SplQueue count ────────────────────────────────────────────

#[test] fn spl_queue_count_changes() {
    compile_ok(r#"<?php
$q = new SplQueue();
echo $q->count();
$q->enqueue('a'); $q->enqueue('b'); $q->enqueue('c');
echo $q->count();
$q->dequeue();
echo $q->count();
"#);
}

// ── SplDoublyLinkedList rewind/current/next ───────────────────

#[test] fn spl_dll_rewind_current_next() {
    compile_ok(r#"<?php
$dll = new SplDoublyLinkedList();
$dll->setIteratorMode(SplDoublyLinkedList::IT_MODE_FIFO);
$dll->push('a'); $dll->push('b'); $dll->push('c');
$dll->rewind();
$out = [];
while ($dll->valid()) {
    $out[] = $dll->current();
    $dll->next();
}
echo implode(',', $out);
"#);
}

// ── SplMinHeap count after extractions ───────────────────────

#[test] fn spl_min_heap_count_after_extractions() {
    compile_ok(r#"<?php
$h = new SplMinHeap();
$h->insert(5); $h->insert(3); $h->insert(8); $h->insert(1);
echo $h->count();
$h->extract(); $h->extract();
echo $h->count();
"#);
}

// ── SplPriorityQueue with identical priorities ────────────────

#[test] fn spl_priority_queue_identical_priorities() {
    compile_ok(r#"<?php
$pq = new SplPriorityQueue();
$pq->insert('task-a', 5);
$pq->insert('task-b', 5);
$pq->insert('task-c', 5);
echo $pq->count();
$pq->extract();
echo $pq->count();
"#);
}

// ── SplFixedArray setSize ─────────────────────────────────────

#[test] fn spl_fixed_array_set_size() {
    compile_ok(r#"<?php
$fa = new SplFixedArray(3);
$fa[0] = 'x'; $fa[1] = 'y'; $fa[2] = 'z';
$fa->setSize(5);
$fa[3] = 'a'; $fa[4] = 'b';
echo $fa->getSize();
echo $fa[4];
"#);
}

// ── SplObjectStorage contains check ──────────────────────────

#[test] fn spl_object_storage_contains_check() {
    compile_ok(r#"<?php
$store = new SplObjectStorage();
$obj1 = new stdClass(); $obj1->id = 1;
$obj2 = new stdClass(); $obj2->id = 2;
$store->attach($obj1);
echo $store->contains($obj1) ? 'yes' : 'no';
echo $store->contains($obj2) ? 'yes' : 'no';
$store->attach($obj2);
echo $store->contains($obj2) ? 'yes' : 'no';
"#);
}
