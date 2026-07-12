use super::helpers::compile_ok;

// ── SplStack (LIFO) ───────────────────────────────────────────

#[test]
fn spl_stack_basic() {
    compile_ok(
        r#"<?php
$stack = new SplStack();
$stack->push(1);
$stack->push(2);
$stack->push(3);
echo $stack->top();
echo $stack->count();
"#,
    );
}

#[test]
fn spl_stack_pop_order() {
    compile_ok(
        r#"<?php
$stack = new SplStack();
foreach ([10, 20, 30] as $v) { $stack->push($v); }
$result = [];
while (!$stack->isEmpty()) { $result[] = $stack->pop(); }
echo implode(',', $result);
"#,
    );
}

#[test]
fn spl_stack_is_empty() {
    compile_ok(
        r#"<?php
$s = new SplStack();
echo $s->isEmpty() ? 'empty' : 'not empty';
$s->push('a');
echo $s->isEmpty() ? 'empty' : 'not empty';
"#,
    );
}

// ── SplQueue (FIFO) ───────────────────────────────────────────

#[test]
fn spl_queue_basic() {
    compile_ok(
        r#"<?php
$q = new SplQueue();
$q->enqueue('first');
$q->enqueue('second');
$q->enqueue('third');
echo $q->count();
"#,
    );
}

#[test]
fn spl_queue_dequeue_order() {
    compile_ok(
        r#"<?php
$q = new SplQueue();
foreach (['a', 'b', 'c'] as $v) { $q->enqueue($v); }
$result = [];
while (!$q->isEmpty()) { $result[] = $q->dequeue(); }
echo implode(',', $result);
"#,
    );
}

#[test]
fn spl_queue_bottom_top() {
    compile_ok(
        r#"<?php
$q = new SplQueue();
$q->enqueue('first');
$q->enqueue('middle');
$q->enqueue('last');
echo $q->bottom() . ',' . $q->top();
"#,
    );
}

// ── SplDoublyLinkedList ───────────────────────────────────────

#[test]
fn spl_dll_push_and_pop() {
    compile_ok(
        r#"<?php
$dll = new SplDoublyLinkedList();
$dll->push('a');
$dll->push('b');
$dll->push('c');
echo $dll->pop();
echo $dll->count();
"#,
    );
}

#[test]
fn spl_dll_unshift_shift() {
    compile_ok(
        r#"<?php
$dll = new SplDoublyLinkedList();
$dll->push('b');
$dll->push('c');
$dll->unshift('a');
echo $dll->shift();
echo $dll->count();
"#,
    );
}

#[test]
fn spl_dll_iterate() {
    compile_ok(
        r#"<?php
$dll = new SplDoublyLinkedList();
$dll->setIteratorMode(SplDoublyLinkedList::IT_MODE_FIFO);
foreach ([1, 2, 3] as $v) { $dll->push($v); }
$result = [];
$dll->rewind();
while ($dll->valid()) { $result[] = $dll->current(); $dll->next(); }
echo implode(',', $result);
"#,
    );
}

// ── SplFixedArray ─────────────────────────────────────────────

#[test]
fn spl_fixed_array_basic() {
    compile_ok(
        r#"<?php
$arr = new SplFixedArray(5);
$arr[0] = 10;
$arr[1] = 20;
$arr[2] = 30;
echo $arr->getSize() . ':' . $arr[1];
"#,
    );
}

#[test]
fn spl_fixed_array_from_array() {
    compile_ok(
        r#"<?php
$arr = SplFixedArray::fromArray([100, 200, 300]);
echo $arr->getSize() . ':' . $arr[2];
"#,
    );
}

#[test]
fn spl_fixed_array_iterate() {
    compile_ok(
        r#"<?php
$arr = SplFixedArray::fromArray([1, 4, 9, 16, 25]);
$sum = 0;
foreach ($arr as $v) { $sum += $v; }
echo $sum;
"#,
    );
}

#[test]
fn spl_fixed_array_resize() {
    compile_ok(
        r#"<?php
$arr = new SplFixedArray(3);
$arr[0] = 'a'; $arr[1] = 'b'; $arr[2] = 'c';
$arr->setSize(5);
$arr[3] = 'd'; $arr[4] = 'e';
echo $arr->getSize() . ':' . $arr[4];
"#,
    );
}

// ── SplMinHeap / SplMaxHeap ───────────────────────────────────

#[test]
fn spl_min_heap_basic() {
    compile_ok(
        r#"<?php
$heap = new SplMinHeap();
$heap->insert(5);
$heap->insert(2);
$heap->insert(8);
$heap->insert(1);
$result = [];
while (!$heap->isEmpty()) { $result[] = $heap->extract(); }
echo implode(',', $result);
"#,
    );
}

#[test]
fn spl_max_heap_basic() {
    compile_ok(
        r#"<?php
$heap = new SplMaxHeap();
$heap->insert(5);
$heap->insert(2);
$heap->insert(8);
$heap->insert(1);
$result = [];
while (!$heap->isEmpty()) { $result[] = $heap->extract(); }
echo implode(',', $result);
"#,
    );
}

#[test]
fn spl_min_heap_top() {
    compile_ok(
        r#"<?php
$heap = new SplMinHeap();
foreach ([30, 10, 50, 20] as $v) { $heap->insert($v); }
echo $heap->top();
"#,
    );
}

// ── SplPriorityQueue ─────────────────────────────────────────

#[test]
fn spl_priority_queue_basic() {
    compile_ok(
        r#"<?php
$pq = new SplPriorityQueue();
$pq->insert('low task',    1);
$pq->insert('high task',   10);
$pq->insert('medium task', 5);
$result = [];
while (!$pq->isEmpty()) { $result[] = $pq->extract(); }
echo implode(',', $result);
"#,
    );
}

#[test]
fn spl_priority_queue_count() {
    compile_ok(
        r#"<?php
$pq = new SplPriorityQueue();
$pq->insert('a', 1);
$pq->insert('b', 2);
$pq->insert('c', 3);
echo $pq->count();
"#,
    );
}

// ── ArrayObject ───────────────────────────────────────────────

#[test]
fn array_object_basic() {
    compile_ok(
        r#"<?php
$ao = new ArrayObject(['x' => 1, 'y' => 2]);
echo $ao['x'];
$ao['z'] = 3;
echo $ao->count();
"#,
    );
}

#[test]
fn array_object_append() {
    compile_ok(
        r#"<?php
$ao = new ArrayObject([1, 2, 3]);
$ao->append(4);
$ao->append(5);
echo $ao->count();
"#,
    );
}

#[test]
fn array_object_iterate() {
    compile_ok(
        r#"<?php
$ao = new ArrayObject(['a' => 1, 'b' => 2, 'c' => 3]);
$sum = 0;
foreach ($ao as $k => $v) { $sum += $v; }
echo $sum;
"#,
    );
}

#[test]
fn array_object_getarraycopy() {
    compile_ok(
        r#"<?php
$ao = new ArrayObject([3, 1, 4, 1, 5]);
$copy = $ao->getArrayCopy();
sort($copy);
echo implode(',', $copy);
"#,
    );
}

#[test]
fn array_object_offsetexists() {
    compile_ok(
        r#"<?php
$ao = new ArrayObject(['name' => 'Alice']);
echo $ao->offsetExists('name') ? 'yes' : 'no';
echo $ao->offsetExists('age')  ? 'yes' : 'no';
"#,
    );
}

// ── ArrayIterator ─────────────────────────────────────────────

#[test]
fn array_iterator_basic() {
    compile_ok(
        r#"<?php
$it = new ArrayIterator([10, 20, 30]);
$sum = 0;
foreach ($it as $v) { $sum += $v; }
echo $sum;
"#,
    );
}

#[test]
fn array_iterator_sort_flags() {
    compile_ok(
        r#"<?php
$it = new ArrayIterator(['banana', 'apple', 'cherry']);
$it->asort();
echo implode(',', iterator_to_array($it));
"#,
    );
}

// ── SplObjectStorage ─────────────────────────────────────────

#[test]
fn spl_object_storage_basic() {
    compile_ok(
        r#"<?php
$store = new SplObjectStorage();
$a = new stdClass(); $a->name = 'Alice';
$b = new stdClass(); $b->name = 'Bob';
$store->attach($a, 'data-a');
$store->attach($b, 'data-b');
echo $store->count();
"#,
    );
}

#[test]
fn spl_object_storage_contains() {
    compile_ok(
        r#"<?php
$store = new SplObjectStorage();
$obj = new stdClass();
echo $store->contains($obj) ? 'yes' : 'no';
$store->attach($obj);
echo $store->contains($obj) ? 'yes' : 'no';
"#,
    );
}

#[test]
fn spl_object_storage_detach() {
    compile_ok(
        r#"<?php
$store = new SplObjectStorage();
$a = new stdClass();
$b = new stdClass();
$store->attach($a);
$store->attach($b);
$store->detach($a);
echo $store->count();
"#,
    );
}

#[test]
fn spl_object_storage_iterate() {
    compile_ok(
        r#"<?php
$store = new SplObjectStorage();
for ($i = 0; $i < 3; $i++) {
    $obj = new stdClass();
    $obj->id = $i;
    $store->attach($obj, "info-$i");
}
$ids = [];
$store->rewind();
while ($store->valid()) {
    $ids[] = $store->getInfo();
    $store->next();
}
sort($ids);
echo implode(',', $ids);
"#,
    );
}
