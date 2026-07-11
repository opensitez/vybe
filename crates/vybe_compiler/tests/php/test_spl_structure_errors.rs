//! SPL structure failure paths only — empty extract, bad offsets, iterator faults.
//! Happy-path push/pop/order tests live in `test_spl.rs` / `test_spl_data_structures.rs`.
//!
//! Correction: empty pop/shift/dequeue/extract/top/bottom throw `RuntimeException`,
//! NOT `UnderflowException` (which is a *subclass* and cannot catch its parent).
//! Verified against the php 8.4 CLI — the earlier `catch (UnderflowException)`
//! expectations would fail on real PHP too.

crate::php_cases! {
    spl_queue_dequeue_on_empty_throws_underflow => {
        r#"<?php
$q = new SplQueue();
try { $q->dequeue(); echo 'ok'; }
catch (RuntimeException $e) { echo 'q-deq'; }
"#,
        ["q-deq"]
    };

    spl_stack_pop_on_empty_throws_underflow => {
        r#"<?php
$s = new SplStack();
try { $s->pop(); echo 'ok'; }
catch (RuntimeException $e) { echo 's-pop'; }
"#,
        ["s-pop"]
    };

    spl_doubly_linked_list_shift_on_empty => {
        r#"<?php
$l = new SplDoublyLinkedList();
try { $l->shift(); echo 'ok'; }
catch (RuntimeException $e) { echo 'dll-shift'; }
"#,
        ["dll-shift"]
    };

    spl_doubly_linked_list_pop_on_empty => {
        r#"<?php
$l = new SplDoublyLinkedList();
try { $l->pop(); echo 'ok'; }
catch (RuntimeException $e) { echo 'dll-pop'; }
"#,
        ["dll-pop"]
    };

    spl_min_heap_extract_on_empty => {
        r#"<?php
$h = new SplMinHeap();
try { $h->extract(); echo 'ok'; }
catch (RuntimeException $e) { echo 'heap-x'; }
"#,
        ["heap-x"]
    };

    spl_priority_queue_extract_on_empty => {
        r#"<?php
$pq = new SplPriorityQueue();
try { $pq->extract(); echo 'ok'; }
catch (RuntimeException $e) { echo 'pq-x'; }
"#,
        ["pq-x"]
    };

    spl_queue_top_on_empty => {
        r#"<?php
$q = new SplQueue();
try { $q->top(); echo 'ok'; }
catch (RuntimeException $e) { echo 'q-top'; }
"#,
        ["q-top"]
    };

    spl_stack_bottom_on_empty => {
        r#"<?php
$s = new SplStack();
try { $s->bottom(); echo 'ok'; }
catch (RuntimeException $e) { echo 's-bot'; }
"#,
        ["s-bot"]
    };

    spl_fixed_array_read_past_end => {
        r#"<?php
$a = new SplFixedArray(2);
try { echo $a[9]; }
catch (OutOfRangeException $e) { echo 'fa-read'; }
"#,
        ["fa-read"]
    };

    spl_fixed_array_write_past_end => {
        r#"<?php
$a = new SplFixedArray(1);
try { $a[4] = 1; echo 'ok'; }
catch (OutOfRangeException $e) { echo 'fa-write'; }
"#,
        ["fa-write"]
    };

    spl_fixed_array_negative_index_read => {
        r#"<?php
$a = SplFixedArray::fromArray([1, 2, 3]);
try { echo $a[-1]; }
catch (OutOfRangeException $e) { echo 'fa-neg'; }
"#,
        ["fa-neg"]
    };

    array_iterator_seek_beyond_last_index => {
        r#"<?php
$it = new ArrayIterator([10, 20]);
try { $it->seek(5); echo $it->current(); }
catch (OutOfBoundsException $e) { echo 'seek-oob'; }
"#,
        ["seek-oob"]
    };

    empty_iterator_current_throws => {
        r#"<?php
$it = new EmptyIterator();
try { $it->current(); echo 'ok'; }
catch (RuntimeException $e) { echo 'empty-cur'; }
"#,
        ["empty-cur"]
    };

    empty_iterator_key_throws => {
        r#"<?php
$it = new EmptyIterator();
try { $it->key(); echo 'ok'; }
catch (RuntimeException $e) { echo 'empty-key'; }
"#,
        ["empty-key"]
    };

    spl_object_storage_offset_get_unattached_returns_null => {
        r#"<?php
$s = new SplObjectStorage();
$o = new stdClass();
echo $s[$o] === null ? 'null' : 'set';
"#,
        ["null"]
    };

    spl_object_storage_detach_unknown_object_is_noop => {
        r#"<?php
$s = new SplObjectStorage();
$o = new stdClass();
$s->detach($o);
echo $s->count();
"#,
        ["0"]
    };

    limit_iterator_negative_max_is_invalid => {
        r#"<?php
$inner = new ArrayIterator([1, 2, 3]);
try { new LimitIterator($inner, 0, -1); echo 'ok'; }
catch (ValueError $e) { echo 'lim-neg'; }
"#,
        ["ok"]
    };

    multiple_iterator_without_attach_has_zero_count => {
        r#"<?php
$m = new MultipleIterator();
echo $m->count();
"#,
        ["0"]
    };

    caching_iterator_out_of_bounds_on_empty_inner => {
        r#"<?php
$inner = new ArrayIterator([]);
$cache = new CachingIterator($inner);
echo $cache->valid() ? 'yes' : 'no';
"#,
        ["no"]
    };

    recursive_iterator_iterator_leaf_has_no_children => {
        r#"<?php
$inner = new RecursiveArrayIterator([1, [2, 3]]);
$rii = new RecursiveIteratorIterator($inner);
$rii->seek(0);
echo $rii->hasChildren() ? 'kids' : 'leaf';
"#,
        ["leaf"]
    };
}
