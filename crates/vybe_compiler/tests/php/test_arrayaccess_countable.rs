use super::helpers::run_prints;

// ── ArrayAccess basic implementation ─────────────────────────

#[test]
fn arrayaccess_offset_exists_and_get() {
    assert_eq!(
        run_prints(
            r#"<?php
class Box implements ArrayAccess {
    private array $data = [];
    public function offsetExists(mixed $k): bool { return isset($this->data[$k]); }
    public function offsetGet(mixed $k): mixed { return $this->data[$k]; }
    public function offsetSet(mixed $k, mixed $v): void { $this->data[$k] = $v; }
    public function offsetUnset(mixed $k): void { unset($this->data[$k]); }
}
$b = new Box;
$b['x'] = 10;
echo $b['x'];
"#
        ),
        vec!["10"]
    );
}
#[test]
fn arrayaccess_isset_delegates_to_offset_exists() {
    assert_eq!(
        run_prints(
            r#"<?php
class Bag implements ArrayAccess {
    private array $d = [];
    public function offsetExists(mixed $k): bool { return array_key_exists($k, $this->d); }
    public function offsetGet(mixed $k): mixed { return $this->d[$k]; }
    public function offsetSet(mixed $k, mixed $v): void { $this->d[$k] = $v; }
    public function offsetUnset(mixed $k): void { unset($this->d[$k]); }
}
$b = new Bag; $b['a'] = 1;
echo isset($b['a']) ? 'yes' : 'no';
echo isset($b['b']) ? 'yes' : 'no';
"#
        ),
        vec!["yesno"]
    );
}
#[test]
fn arrayaccess_unset_removes_key() {
    assert_eq!(
        run_prints(
            r#"<?php
class Store implements ArrayAccess {
    private array $d = ['k' => 'v'];
    public function offsetExists(mixed $k): bool { return isset($this->d[$k]); }
    public function offsetGet(mixed $k): mixed { return $this->d[$k]; }
    public function offsetSet(mixed $k, mixed $v): void { $this->d[$k] = $v; }
    public function offsetUnset(mixed $k): void { unset($this->d[$k]); }
}
$s = new Store;
unset($s['k']);
echo isset($s['k']) ? 'exists' : 'gone';
"#
        ),
        vec!["gone"]
    );
}
#[test]
fn arrayaccess_push_with_null_key() {
    assert_eq!(
        run_prints(
            r#"<?php
class List2 implements ArrayAccess {
    private array $d = [];
    public function offsetExists(mixed $k): bool { return isset($this->d[$k]); }
    public function offsetGet(mixed $k): mixed { return $this->d[$k]; }
    public function offsetSet(mixed $k, mixed $v): void {
        if ($k === null) $this->d[] = $v;
        else $this->d[$k] = $v;
    }
    public function offsetUnset(mixed $k): void { unset($this->d[$k]); }
    public function toArray(): array { return $this->d; }
}
$l = new List2;
$l[] = 'a'; $l[] = 'b'; $l[] = 'c';
echo implode(',', $l->toArray());
"#
        ),
        vec!["a,b,c"]
    );
}

// ── ArrayAccess combined with Countable ───────────────────────

#[test]
fn countable_interface_count() {
    assert_eq!(
        run_prints(
            r#"<?php
class Counter implements Countable {
    private array $items;
    public function __construct(array $items) { $this->items = $items; }
    public function count(): int { return count($this->items); }
}
$c = new Counter([1,2,3,4,5]);
echo count($c);
"#
        ),
        vec!["5"]
    );
}
#[test]
fn arrayaccess_and_countable_combined() {
    assert_eq!(
        run_prints(
            r#"<?php
class Collection implements ArrayAccess, Countable {
    private array $d = [];
    public function offsetExists(mixed $k): bool { return isset($this->d[$k]); }
    public function offsetGet(mixed $k): mixed { return $this->d[$k]; }
    public function offsetSet(mixed $k, mixed $v): void { $this->d[$k] = $v; }
    public function offsetUnset(mixed $k): void { unset($this->d[$k]); }
    public function count(): int { return count($this->d); }
}
$c = new Collection;
$c['a'] = 1; $c['b'] = 2; $c['c'] = 3;
echo count($c);
"#
        ),
        vec!["3"]
    );
}

// ── IteratorAggregate ─────────────────────────────────────────

#[test]
fn iterator_aggregate_foreach() {
    assert_eq!(
        run_prints(
            r#"<?php
class NumberRange implements IteratorAggregate {
    public function __construct(private int $from, private int $to) {}
    public function getIterator(): ArrayIterator {
        return new ArrayIterator(range($this->from, $this->to));
    }
}
$r = new NumberRange(1, 4);
foreach ($r as $n) echo $n;
"#
        ),
        vec!["1234"]
    );
}
#[test]
fn iterator_rewind_and_reuse() {
    assert_eq!(
        run_prints(
            r#"<?php
class Words implements IteratorAggregate {
    public function getIterator(): ArrayIterator {
        return new ArrayIterator(['foo','bar','baz']);
    }
}
$w = new Words;
foreach ($w as $v) echo $v[0];
foreach ($w as $v) echo strtoupper($v[0]);
"#
        ),
        vec!["fbbFBB"]
    );
}

// ── Iterator interface ────────────────────────────────────────

#[test]
fn iterator_interface_manual() {
    assert_eq!(
        run_prints(
            r#"<?php
class Countdown implements Iterator {
    private int $cur;
    public function __construct(private int $start) { $this->cur = $start; }
    public function current(): int { return $this->cur; }
    public function key(): int { return $this->start - $this->cur; }
    public function next(): void { $this->cur--; }
    public function rewind(): void { $this->cur = $this->start; }
    public function valid(): bool { return $this->cur > 0; }
}
foreach (new Countdown(3) as $n) echo $n;
"#
        ),
        vec!["321"]
    );
}
#[test]
fn generator_is_iterator() {
    assert_eq!(
        run_prints(
            r#"<?php
function evens(int $max): Generator {
    for ($i = 2; $i <= $max; $i += 2) yield $i;
}
$sum = 0;
foreach (evens(10) as $n) $sum += $n;
echo $sum;
"#
        ),
        vec!["30"]
    );
}

// ── ArrayObject ───────────────────────────────────────────────

#[test]
fn array_object_access_and_append() {
    assert_eq!(
        run_prints(
            r#"<?php
$ao = new ArrayObject(['x' => 1]);
$ao['y'] = 2;
$ao->append(3);
echo $ao['x'] . ',' . $ao['y'] . ',' . $ao->count();
"#
        ),
        vec!["1,2,3"]
    );
}
#[test]
fn array_object_iterate() {
    assert_eq!(
        run_prints(
            r#"<?php
$ao = new ArrayObject(['a' => 1, 'b' => 2, 'c' => 3]);
foreach ($ao as $k => $v) echo $k . $v;
"#
        ),
        vec!["a1b2c3"]
    );
}
#[test]
fn array_object_sort() {
    assert_eq!(
        run_prints(
            r#"<?php
$ao = new ArrayObject([3,1,2]);
$ao->asort();
echo implode(',', (array)$ao);
"#
        ),
        vec!["1,2,3"]
    );
}
