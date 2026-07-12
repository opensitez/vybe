//! SPL data structures and iterators — runtime happy paths.

crate::php_cases! {
    spl_autoload_register_callable => {
        r#"<?php
spl_autoload_register(function ($c) {});
echo class_exists('stdClass') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    spl_object_hash_unique => {
        r#"<?php
echo strlen(spl_object_hash(new stdClass())) > 0 ? 'hash' : 'no';
"#,
        ["hash"]
    };

    spl_classes_returns_array => {
        r#"<?php
echo is_array(spl_classes()) ? 'arr' : 'no';
"#,
        ["arr"]
    };

    spl_priority_queue_insert_extract => {
        r#"<?php
$q = new SplPriorityQueue();
$q->insert('a', 1);
$q->insert('b', 3);
echo $q->extract();
"#,
        ["b"]
    };

    spl_stack_push_pop => {
        r#"<?php
$s = new SplStack();
$s->push(1);
$s->push(2);
echo $s->pop();
"#,
        ["2"]
    };

    spl_queue_enqueue_dequeue => {
        r#"<?php
$q = new SplQueue();
$q->enqueue(1);
$q->enqueue(2);
echo $q->dequeue();
"#,
        ["1"]
    };

    spl_doubly_linked_list_push_shift => {
        r#"<?php
$l = new SplDoublyLinkedList();
$l->push('x');
$l->push('y');
echo $l->shift();
"#,
        ["x"]
    };

    spl_fixed_array_set_get => {
        r#"<?php
$a = new SplFixedArray(2);
$a[0] = 'a';
echo $a[0];
"#,
        ["a"]
    };

    spl_heap_max_extract => {
        r#"<?php
class MaxHeap extends SplMaxHeap {
    protected function compare($a, $b): int { return $a <=> $b; }
}
$h = new MaxHeap();
$h->insert(1);
$h->insert(3);
echo $h->extract();
"#,
        ["3"]
    };

    spl_min_heap_extract => {
        r#"<?php
class MinHeap extends SplMinHeap {
    protected function compare($a, $b): int { return $a <=> $b; }
}
$h = new MinHeap();
$h->insert(5);
$h->insert(2);
echo $h->extract();
"#,
        ["5"]
    };

    spl_array_object_offset_set => {
        r#"<?php
$o = new ArrayObject(['k' => 'v']);
echo $o['k'];
"#,
        ["v"]
    };

    spl_array_iterator_foreach => {
        r#"<?php
$it = new ArrayIterator([1, 2]);
$s = 0;
foreach ($it as $n) { $s += $n; }
echo $s;
"#,
        ["3"]
    };

    spl_file_object_read_line => {
        r#"<?php
$f = new SplFileObject('php://memory', 'r+');
$f->fwrite("line\n");
$f->rewind();
echo trim($f->fgets());
"#,
        ["line"]
    };

    spl_temp_fileobject_write => {
        r#"<?php
$f = new SplTempFileObject();
$f->fwrite('tmp');
$f->rewind();
echo $f->fgets();
"#,
        ["tmp"]
    };

    spl_object_storage_attach_offset => {
        r#"<?php
$s = new SplObjectStorage();
$o = new stdClass();
$s->attach($o, 'meta');
echo $s[$o];
"#,
        ["meta"]
    };

    spl_object_storage_count => {
        r#"<?php
$s = new SplObjectStorage();
$s->attach(new stdClass());
echo count($s);
"#,
        ["1"]
    };

    spl_observer_spl_subject => {
        r#"<?php
class W implements SplObserver {
    public int $n = 0;
    public function update(SplSubject $s): void { $this->n++; }
}
class S implements SplSubject {
    private array $o = [];
    public function attach(SplObserver $o): void { $this->o[] = $o; }
    public function detach(SplObserver $o): void {}
    public function notify(): void { foreach ($this->o as $o) { $o->update($this); } }
}
$w = new W();
$s = new S();
$s->attach($w);
$s->notify();
echo $w->n;
"#,
        ["1"]
    };

    spl_heap_count_after_insert => {
        r#"<?php
class MaxHeap extends SplMaxHeap {
    protected function compare($a, $b): int { return $a <=> $b; }
}
$h = new MaxHeap();
$h->insert(1);
$h->insert(2);
echo $h->count();
"#,
        ["2"]
    };

    spl_iterator_to_array => {
        r#"<?php
echo count(iterator_to_array(new ArrayIterator([1, 2, 3])));
"#,
        ["3"]
    };

    spl_callback_filter_iterator => {
        r#"<?php
$it = new CallbackFilterIterator(new ArrayIterator([1, 2, 3]), fn($n) => $n > 1);
echo count(iterator_to_array($it));
"#,
        ["2"]
    };

    spl_limit_iterator => {
        r#"<?php
$it = new LimitIterator(new ArrayIterator([1, 2, 3, 4]), 1, 2);
echo implode('', iterator_to_array($it));
"#,
        ["23"]
    };

    spl_append_iterator => {
        r#"<?php
$a = new ArrayIterator([1]);
$b = new ArrayIterator([2]);
$app = new AppendIterator();
$app->append($a);
$app->append($b);
echo count(iterator_to_array($app));
"#,
        ["1"]
    };

    spl_no_rewind_iterator => {
        r#"<?php
$base = new ArrayIterator([1, 2]);
$it = new NoRewindIterator($base);
echo $it->valid() ? 'valid' : 'no';
"#,
        ["valid"]
    };

    spl_empty_iterator => {
        r#"<?php
$it = new EmptyIterator();
echo $it->valid() ? 'yes' : 'no';
"#,
        ["no"]
    };

    spl_infinite_iterator => {
        r#"<?php
$it = new InfiniteIterator(new ArrayIterator([1]));
$it->rewind();
echo $it->current();
"#,
        ["1"]
    };

    spl_caching_iterator_string => {
        r#"<?php
$it = new CachingIterator(new ArrayIterator(['a']));
$it->next();
echo $it->getCache();
"#,
        ["a"]
    };

    spl_recursive_tree_iterator => {
        r#"<?php
$it = new RecursiveTreeIterator(new RecursiveArrayIterator(['a' => ['b']]));
$it->rewind();
echo $it->getPrefix();
"#,
        ["\\-"]
    };

    spl_recursive_iterator_iterator => {
        r#"<?php
$tree = new RecursiveArrayIterator([1, [2, 3]]);
$it = new RecursiveIteratorIterator($tree);
echo array_sum(iterator_to_array($it));
"#,
        ["5"]
    };

    spl_filter_iterator => {
        r#"<?php
$it = new FilterIterator(new ArrayIterator([1, 2, 3]));
echo $it->valid() ? 'ok' : 'no';
"#,
        ["ok"]
    };

    spl_regex_iterator => {
        r#"<?php
$it = new RegexIterator(new ArrayIterator(['a1', 'b2']), '/\d/');
echo count(iterator_to_array($it));
"#,
        ["2"]
    };

    spl_multiple_iterator => {
        r#"<?php
$m = new MultipleIterator();
$m->attachIterator(new ArrayIterator([1]));
$m->attachIterator(new ArrayIterator([2]));
$m->rewind();
echo $m->valid() ? 'yes' : 'no';
"#,
        ["yes"]
    };

    spl_iterator_iterator_wrap => {
        r#"<?php
$gen = (function () { yield 1; yield 2; })();
$it = new IteratorIterator($gen);
echo implode('', iterator_to_array($it));
"#,
        ["12"]
    };

    spl_array_object_count => {
        r#"<?php
$o = new ArrayObject([1, 2, 3]);
echo count($o);
"#,
        ["3"]
    };

    spl_fixed_array_count => {
        r#"<?php
echo (new SplFixedArray(5))->count();
"#,
        ["5"]
    };

    spl_stack_count => {
        r#"<?php
$s = new SplStack();
$s->push(1);
$s->push(2);
echo $s->count();
"#,
        ["2"]
    };
}
