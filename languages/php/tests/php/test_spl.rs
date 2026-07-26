use super::helpers::{compile_ok, run_prints};

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

#[test]
fn spl_priority_queue_extract_both() {
    compile_ok(
        r#"<?php
$pq = new SplPriorityQueue();
$pq->setExtractFlags(SplPriorityQueue::EXTR_BOTH);
$pq->insert('low', 1);
$pq->insert('high', 9);
$pq->insert('mid', 5);
$item = $pq->extract();
echo $item['data'];
echo $item['priority'];
"#,
    );
}

#[test]
fn spl_priority_queue_extract_data_only() {
    compile_ok(
        r#"<?php
$pq = new SplPriorityQueue();
$pq->setExtractFlags(SplPriorityQueue::EXTR_DATA);
$pq->insert('first', 1);
$pq->insert('second', 10);
echo $pq->current();
echo $pq->key();
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

#[test]
fn array_object_iterator_mode() {
    compile_ok(
        r#"<?php
$ao = new ArrayObject(['a' => 1, 'b' => 2], ArrayObject::STD_PROP_LIST);
$ao->ksort();
$it = $ao->getIterator();
foreach ($it as $k => $v) { echo $k . $v; }
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

#[test]
fn array_iterator_compatibility() {
    compile_ok(
        r#"<?php
$it = new ArrayIterator([3, 1, 2]);
$it->asort(SORT_NUMERIC);
echo $it->count();
echo $it->current();
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
fn spl_object_storage_info_api() {
    compile_ok(
        r#"<?php
$store = new SplObjectStorage();
$obj = new stdClass();
$store->attach($obj, ['tag' => 'alpha']);
echo $store->contains($obj) ? 'yes' : 'no';
echo $store->getInfo()['tag'];
echo $store->getHash($obj);
"#,
    );
}

#[test]
fn spl_object_storage_replace_info() {
    compile_ok(
        r#"<?php
$store = new SplObjectStorage();
$obj = new stdClass();
$store->attach($obj, 'first');
$store->rewind();
$store->setInfo('second');
echo $store->getInfo();
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

#[test]
fn spl_stack_pop_and_count_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$stack = new SplStack();
$stack->push('a');
$stack->push('b');
$stack->push('c');
echo $stack->count();
echo '|';
echo $stack->pop();
echo '|';
echo $stack->count();
"#,
        ),
        &["3|c|2"]
    );
}

#[test]
fn spl_queue_fifo_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$queue = new SplQueue();
$queue->setIteratorMode(SplDoublyLinkedList::IT_MODE_FIFO | SplDoublyLinkedList::IT_MODE_KEEP);
$queue->enqueue(1);
$queue->enqueue(2);
$queue->enqueue(3);
$first = $queue->shift();
$last = $queue->top();
echo $first . '|' . $last . '|' . $queue->count();
"#,
        ),
        &["1|3|2"]
    );
}

#[test]
fn spl_fixed_array_indexes_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$fixed = new SplFixedArray(4);
$fixed[1] = 10;
$fixed[2] = 20;
echo $fixed->getSize();
echo '|';
echo $fixed->count();
echo '|';
echo $fixed[1];
"#,
        ),
        &["4|4|10"]
    );
}

#[test]
fn spl_priority_queue_pairs_with_both_extract_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$pq = new SplPriorityQueue();
$pq->setExtractFlags(SplPriorityQueue::EXTR_BOTH);
$pq->insert('low', 1);
$pq->insert('high', 9);
$pq->insert('mid', 5);
$item = $pq->extract();
echo $item['data'];
echo '|';
echo $item['priority'];
echo '|';
echo $pq->count();
"#,
        ),
        &["high|9|2"]
    );
}

#[test]
fn array_object_std_to_array_copy_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$obj = new ArrayObject(['x' => 1, 'y' => 2]);
$snapshot = $obj->getArrayCopy();
$obj['x'] = 9;
echo $snapshot['x'];
echo '|';
echo $obj->count();
"#,
        ),
        &["1|2"]
    );
}

#[test]
fn array_iterator_seek_and_key_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$it = new ArrayIterator(['a' => 1, 'b' => 2, 'c' => 3]);
$it->seek(2);
echo $it->key();
echo '|';
echo $it->current();
"#,
        ),
        &["2|3"]
    );
}

#[test]
fn spl_object_storage_attach_info_and_get_info_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$store = new SplObjectStorage();
$obj = new stdClass();
$store->attach($obj, ['role' => 'admin']);
echo $store->contains($obj) ? 'yes' : 'no';
echo '|';
echo $store->getInfo()['role'];
"#,
        ),
        &["yes|admin"]
    );
}

#[test]
fn spl_object_storage_keyed_iteration_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$store = new SplObjectStorage();
for ($i = 0; $i < 3; $i++) {
    $obj = new stdClass();
    $obj->i = $i;
    $store->attach($obj, $i);
}
$seen = [];
for ($store->rewind(); $store->valid(); $store->next()) {
    $seen[] = $store->key();
}
sort($seen);
echo implode(',', $seen);
"#,
        ),
        &["0,1,2"]
    );
}

#[test]
fn spl_dll_lifo_iterator_mode_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$dll = new SplDoublyLinkedList();
$dll->setIteratorMode(SplDoublyLinkedList::IT_MODE_LIFO | SplDoublyLinkedList::IT_MODE_KEEP);
$dll->push(1);
$dll->push(2);
$dll->push(3);
$it = [];
$dll->rewind();
while ($dll->valid()) {
    $it[] = $dll->current();
    $dll->next();
}
echo implode(',', $it);
"#,
        ),
        &["3,2,1"]
    );
}

#[test]
fn spl_stack_iterator_mode_fifo_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$stack = new SplStack();
$stack->setIteratorMode(SplDoublyLinkedList::IT_MODE_FIFO | SplDoublyLinkedList::IT_MODE_KEEP);
$stack->push(10);
$stack->push(20);
$stack->push(30);
$stack->rewind();
$out = [];
while ($stack->valid()) {
    $out[] = $stack->current();
    $stack->next();
}
echo implode('|', $out);
        "#,
        ),
        &["10|20|30"]
    );
}

#[test]
fn spl_fixed_array_offset_unset_and_reassign_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$arr = new SplFixedArray(3);
$arr[0] = 'first';
$arr[1] = 'second';
$arr[2] = 'third';
unset($arr[1]);
$arr[1] = 'rebound';
echo $arr->count();
echo '|';
echo $arr[1];
"#,
        ),
        &["3|rebound"]
    );
}

#[test]
fn spl_priority_queue_extract_flags_extr_data_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$pq = new SplPriorityQueue();
$pq->setExtractFlags(SplPriorityQueue::EXTR_DATA);
$pq->insert('low', 1);
$pq->insert('high', 10);
$pq->insert('mid', 5);
echo $pq->current();
echo '|';
echo $pq->extract();
echo '|';
echo $pq->count();
"#,
        ),
        &["high|high|2"]
    );
}

#[test]
fn spl_object_storage_info_persistence_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$store = new SplObjectStorage();
$obj = new stdClass();
$store->attach($obj, ['role' => 'member']);
$store->rewind();
$info = $store->getInfo();
$info['role'] = 'admin';
$store->setInfo($info);
echo $store->getInfo()['role'];
$store->detach($obj);
echo '|';
echo $store->contains($obj) ? 'present' : 'gone';
"#,
        ),
        &["admin|gone"]
    );
}

#[test]
fn array_object_array_access_modes_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
    $ao = new ArrayObject([1, 2, 3], ArrayObject::ARRAY_AS_PROPS);
    unset($ao[1]);
    $ao->append(4);
    echo $ao->count();
echo '|';
$ao['2'] = 9;
echo $ao[2];
"#,
        ),
        &["3|9"]
    );
}

#[test]
fn array_iterator_asort_numeric_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$it = new ArrayIterator([3, 1, 2], ArrayIterator::STD_PROP_LIST);
$it->asort();
$out = [];
for ($it->rewind(); $it->valid(); $it->next()) {
    $out[] = $it->current();
}
echo implode('-', $out);
"#,
        ),
        &["1-2-3"]
    );
}

#[test]
fn spl_stack_iterator_mode_lifo_keep_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$stack = new SplStack();
$stack->setIteratorMode(SplDoublyLinkedList::IT_MODE_LIFO | SplDoublyLinkedList::IT_MODE_KEEP);
$stack->push('first');
$stack->push('second');
$stack->push('third');
$iterated = [];
$stack->rewind();
while ($stack->valid()) {
    $iterated[] = $stack->current();
    $stack->next();
}
echo implode(':', $iterated);
"#,
        ),
        &["third:second:first"]
    );
}

#[test]
fn spl_queue_rewind_and_count_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$queue = new SplQueue();
$queue->setIteratorMode(SplDoublyLinkedList::IT_MODE_FIFO | SplDoublyLinkedList::IT_MODE_KEEP);
$queue->enqueue(1);
$queue->enqueue(2);
$queue->enqueue(3);
$queue->rewind();
echo $queue->count();
echo '|';
echo $queue->current();
echo '|';
$queue->next();
echo $queue->current();
"#,
        ),
        &["3|1|2"]
    );
}

#[test]
fn spl_heap_extract_and_top_after_extraction_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$heap = new SplMinHeap();
$heap->insert(9);
$heap->insert(4);
$heap->insert(7);
$heap->extract();
echo $heap->count();
echo '|';
echo $heap->top();
"#,
        ),
        &["2|7"]
    );
}

#[test]
fn spl_priority_queue_extract_with_data_and_priority_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$pq = new SplPriorityQueue();
$pq->setExtractFlags(SplPriorityQueue::EXTR_BOTH);
$pq->insert('low', 1);
$pq->insert('mid', 2);
$pq->insert('high', 3);
$next = $pq->extract();
echo $next['data'];
echo ':';
echo $next['priority'];
echo '|';
echo $pq->count();
"#,
        ),
        &["high:3|2"]
    );
}

#[test]
fn spl_priority_queue_extract_priority_only_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$pq = new SplPriorityQueue();
$pq->setExtractFlags(SplPriorityQueue::EXTR_PRIORITY);
$pq->insert('alpha', 5);
$pq->insert('beta', 10);
echo $pq->extract();
echo '|';
echo $pq->count();
"#,
        ),
        &["10|1"]
    );
}

#[test]
fn spl_priority_queue_custom_compare_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class ScoreQueue extends SplPriorityQueue {
    public function compare($a, $b) {
        return $a <=> $b;
    }
}
$pq = new ScoreQueue();
$pq->setExtractFlags(SplPriorityQueue::EXTR_BOTH);
$pq->insert('low', 1);
$pq->insert('high', 9);
$pq->insert('mid', 5);
$first = $pq->extract()['data'];
$second = $pq->extract()['data'];
echo $first;
echo '|';
echo $second;
"#,
        ),
        &["high|mid"]
    );
}

#[test]
fn spl_fixed_array_rewind_after_set_size_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$array = new SplFixedArray(2);
$array[0] = 'a';
$array[1] = 'b';
$array->setSize(4);
$array[2] = 'c';
$array[3] = 'd';
echo $array->count();
echo '|';
echo $array->offsetGet(2);
"#,
        ),
        &["4|c"]
    );
}

#[test]
fn spl_doubly_linked_list_delete_mode_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$dll = new SplDoublyLinkedList();
$dll->setIteratorMode(SplDoublyLinkedList::IT_MODE_FIFO | SplDoublyLinkedList::IT_MODE_DELETE);
$dll->push(1);
$dll->push(2);
$dll->push(3);
$dll->rewind();
echo $dll->current();
echo '|';
$dll->next();
echo $dll->current();
"#,
        ),
        &["1|2"]
    );
}

#[test]
fn array_object_array_as_props_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$obj = new ArrayObject([], ArrayObject::ARRAY_AS_PROPS);
$obj->alpha = 1;
$obj['beta'] = 2;
echo $obj->alpha;
echo '|';
echo $obj['beta'];
echo '|';
echo $obj->count();
"#,
        ),
        &["1|2|2"]
    );
}

#[test]
fn array_iterator_filter_mode_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$it = new ArrayIterator(['a' => 3, 'b' => 1, 'c' => 2]);
$it->natcasesort();
$result = [];
for ($it->rewind(); $it->valid(); $it->next()) {
    $result[] = $it->current();
}
echo implode('|', $result);
"#,
        ),
        &["1|2|3"]
    );
}

#[test]
fn spl_object_storage_reusing_object_reference_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$store = new SplObjectStorage();
$obj = new stdClass();
$store->attach($obj, 'first');
$store[$obj] = 'manual';
echo $store->count();
echo '|';
echo $store->offsetGet($obj);
echo '|';
echo $store->getInfo();
"#,
        ),
        &["1|manual|first"]
    );
}
