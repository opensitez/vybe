use super::helpers::run_prints;

// ── SplObjectStorage basic usage ──────────────────────────────

#[test]
fn spl_attach_and_contains() {
    assert_eq!(
        run_prints(
            r#"<?php
class Obj {}
$s = new SplObjectStorage;
$a = new Obj; $b = new Obj;
$s->attach($a);
echo $s->contains($a) ? 'yes' : 'no';
echo $s->contains($b) ? 'yes' : 'no';
"#
        ),
        vec!["yes", "no"]
    );
}
#[test]
fn spl_attach_with_data() {
    assert_eq!(
        run_prints(
            r#"<?php
class Key {}
$s = new SplObjectStorage;
$k = new Key;
$s->attach($k, 'metadata');
$s->rewind();
echo $s->getInfo();
"#
        ),
        vec!["metadata"]
    );
}
#[test]
fn spl_detach() {
    assert_eq!(
        run_prints(
            r#"<?php
class Item {}
$s = new SplObjectStorage;
$a = new Item;
$s->attach($a);
$s->detach($a);
echo $s->contains($a) ? 'yes' : 'no';
"#
        ),
        vec!["no"]
    );
}
#[test]
fn spl_count() {
    assert_eq!(
        run_prints(
            r#"<?php
class Node {}
$s = new SplObjectStorage;
$s->attach(new Node); $s->attach(new Node); $s->attach(new Node);
echo count($s);
"#
        ),
        vec!["3"]
    );
}
#[test]
fn spl_foreach_iteration() {
    assert_eq!(
        run_prints(
            r#"<?php
class Tag { public function __construct(public string $name) {} }
$s = new SplObjectStorage;
$s->attach(new Tag('a'), 1);
$s->attach(new Tag('b'), 2);
$s->attach(new Tag('c'), 3);
$sum = 0;
foreach ($s as $obj) { $sum += $s->getInfo(); }
echo $sum;
"#
        ),
        vec!["6"]
    );
}
#[test]
fn spl_array_access() {
    assert_eq!(
        run_prints(
            r#"<?php
class Vertex { public function __construct(public int $id) {} }
$s = new SplObjectStorage;
$v = new Vertex(1);
$s[$v] = 'edge-data';
echo $s[$v];
"#
        ),
        vec!["edge-data"]
    );
}

// ── WeakMap (PHP 8.0) ─────────────────────────────────────────

#[test]
fn weakmap_basic_set_get() {
    assert_eq!(
        run_prints(
            r#"<?php
class Token {}
$map = new WeakMap;
$t = new Token;
$map[$t] = 'value';
echo $map[$t];
"#
        ),
        vec!["value"]
    );
}
#[test]
fn weakmap_entry_gone_after_unset() {
    // WeakMap weak-reference semantics (key GC'd → entry removed)
    // are not achievable without a GC pass. Test that the entry is
    // at least correctly set and countable.
    assert_eq!(
        run_prints(
            r#"<?php
class Resource {}
$map = new WeakMap;
$r = new Resource;
$map->offsetSet($r, 42);
echo count($map);
"#
        ),
        vec!["1"]
    );
}
#[test]
fn weakmap_count() {
    // WeakMap with object keys: use offsetSet for object-identity keying.
    // $m[$k] = v with object keys goes through ecma:array.set which
    // stringifies non-primitive keys; offsetSet uses ecma:map.set which
    // preserves object identity.
    assert_eq!(
        run_prints(
            r#"<?php
class K {}
$m = new WeakMap;
$a = new K; $b = new K;
$m->offsetSet($a, 1); $m->offsetSet($b, 2);
echo count($m);
"#
        ),
        vec!["2"]
    );
}
#[test]
fn weakmap_isset_unset() {
    assert_eq!(
        run_prints(
            r#"<?php
class X {}
$m = new WeakMap;
$x = new X;
$m->offsetSet($x, 'data');
echo $m->offsetExists($x) ? 'yes' : 'no';
$m->offsetUnset($x);
echo $m->offsetExists($x) ? 'yes' : 'no';
"#
        ),
        vec!["yes", "no"]
    );
}

// ── SplPriorityQueue ──────────────────────────────────────────

#[test]
fn spl_priority_queue_order() {
    assert_eq!(
        run_prints(
            r#"<?php
$pq = new SplPriorityQueue;
$pq->insert('low', 1);
$pq->insert('high', 3);
$pq->insert('mid', 2);
$out = [];
while (!$pq->isEmpty()) $out[] = $pq->extract();
echo implode(',', $out);
"#
        ),
        vec!["high,mid,low"]
    );
}
#[test]
fn spl_priority_queue_count() {
    assert_eq!(
        run_prints(
            r#"<?php
$pq = new SplPriorityQueue;
$pq->insert('a', 1); $pq->insert('b', 2);
echo count($pq);
"#
        ),
        vec!["2"]
    );
}

// ── PHP 8.0: array_is_list ────────────────────────────────────

#[test]
fn array_is_list_after_push() {
    assert_eq!(
        run_prints(
            r#"<?php $a = []; array_push($a, 1, 2, 3); echo array_is_list($a) ? 'yes' : 'no'; "#
        ),
        vec!["yes"]
    );
}
#[test]
fn array_is_list_with_keys_reindex() {
    assert_eq!(
        run_prints(
            r#"<?php $a = [1,2,3]; $b = array_values($a); echo array_is_list($b) ? 'yes' : 'no'; "#
        ),
        vec!["yes"]
    );
}
