use super::helpers::compile_ok;

// ── Iterator interface ────────────────────────────────────────

#[test]
fn iterator_basic() {
    compile_ok(
        r#"<?php
class NumberRange implements Iterator {
    private int $current;
    public function __construct(
        private int $start,
        private int $end
    ) { $this->current = $start; }
    public function current(): int { return $this->current; }
    public function key(): int { return $this->current - $this->start; }
    public function next(): void { $this->current++; }
    public function rewind(): void { $this->current = $this->start; }
    public function valid(): bool { return $this->current <= $this->end; }
}
$range = new NumberRange(1, 5);
foreach ($range as $k => $v) { echo "$k:$v "; }
"#,
    );
}

#[test]
fn iterator_reusable() {
    compile_ok(
        r#"<?php
class Letters implements Iterator {
    private int $pos = 0;
    private array $letters = ['a', 'b', 'c'];
    public function current(): string { return $this->letters[$this->pos]; }
    public function key(): int { return $this->pos; }
    public function next(): void { $this->pos++; }
    public function rewind(): void { $this->pos = 0; }
    public function valid(): bool { return $this->pos < count($this->letters); }
}
$it = new Letters();
foreach ($it as $v) { echo $v; }
foreach ($it as $v) { echo $v; }
"#,
    );
}

#[test]
fn iterator_manual_control() {
    compile_ok(
        r#"<?php
class Counter implements Iterator {
    private int $i = 0;
    public function __construct(private int $max) {}
    public function current(): int  { return $this->i; }
    public function key(): int      { return $this->i; }
    public function next(): void    { $this->i++; }
    public function rewind(): void  { $this->i = 0; }
    public function valid(): bool   { return $this->i < $this->max; }
}
$c = new Counter(3);
$c->rewind();
while ($c->valid()) {
    echo $c->current() . ' ';
    $c->next();
}
"#,
    );
}

#[test]
fn iterator_infinite_with_limit() {
    compile_ok(
        r#"<?php
class Fibonacci implements Iterator {
    private int $a = 0, $b = 1, $step = 0;
    public function current(): int  { return $this->a; }
    public function key(): int      { return $this->step; }
    public function next(): void    { [$this->a, $this->b] = [$this->b, $this->a + $this->b]; $this->step++; }
    public function rewind(): void  { $this->a = 0; $this->b = 1; $this->step = 0; }
    public function valid(): bool   { return true; }
}
$fib = new Fibonacci();
$result = [];
$fib->rewind();
for ($i = 0; $i < 8; $i++) { $result[] = $fib->current(); $fib->next(); }
echo implode(',', $result);
"#,
    );
}

// ── IteratorAggregate ─────────────────────────────────────────

#[test]
fn iterator_aggregate_basic() {
    compile_ok(
        r#"<?php
class Collection implements IteratorAggregate {
    private array $items = [];
    public function add(mixed $item): void { $this->items[] = $item; }
    public function getIterator(): ArrayIterator { return new ArrayIterator($this->items); }
}
$c = new Collection();
$c->add('a'); $c->add('b'); $c->add('c');
foreach ($c as $item) { echo $item; }
"#,
    );
}

#[test]
fn iterator_aggregate_wrapped() {
    compile_ok(
        r#"<?php
class FilteredCollection implements IteratorAggregate {
    public function __construct(private array $items, private callable $filter) {}
    public function getIterator(): ArrayIterator {
        return new ArrayIterator(array_values(array_filter($this->items, $this->filter)));
    }
}
$evens = new FilteredCollection([1, 2, 3, 4, 5, 6], fn($n) => $n % 2 === 0);
foreach ($evens as $v) { echo $v . ' '; }
"#,
    );
}

// ── ArrayAccess interface ─────────────────────────────────────

#[test]
fn array_access_basic() {
    compile_ok(
        r#"<?php
class TypedCollection implements ArrayAccess {
    private array $data = [];
    public function offsetExists(mixed $offset): bool { return isset($this->data[$offset]); }
    public function offsetGet(mixed $offset): mixed   { return $this->data[$offset] ?? null; }
    public function offsetSet(mixed $offset, mixed $value): void {
        if ($offset === null) { $this->data[] = $value; }
        else { $this->data[$offset] = $value; }
    }
    public function offsetUnset(mixed $offset): void { unset($this->data[$offset]); }
}
$c = new TypedCollection();
$c[] = 'first';
$c[] = 'second';
$c['named'] = 'third';
echo $c[0] . ',' . $c['named'];
"#,
    );
}

#[test]
fn array_access_exists_unset() {
    compile_ok(
        r#"<?php
class Registry implements ArrayAccess {
    private array $store = [];
    public function offsetExists(mixed $k): bool  { return array_key_exists($k, $this->store); }
    public function offsetGet(mixed $k): mixed    { return $this->store[$k]; }
    public function offsetSet(mixed $k, mixed $v): void { $this->store[$k] = $v; }
    public function offsetUnset(mixed $k): void   { unset($this->store[$k]); }
}
$r = new Registry();
$r['key'] = 'value';
echo isset($r['key']) ? 'exists' : 'missing';
unset($r['key']);
echo isset($r['key']) ? 'exists' : 'missing';
"#,
    );
}

// ── Countable interface ───────────────────────────────────────

#[test]
fn countable_basic() {
    compile_ok(
        r#"<?php
class WordList implements Countable {
    private array $words = [];
    public function add(string $w): void { $this->words[] = $w; }
    public function count(): int { return count($this->words); }
}
$wl = new WordList();
$wl->add('hello'); $wl->add('world');
echo count($wl);
"#,
    );
}

#[test]
fn countable_with_array_access() {
    compile_ok(
        r#"<?php
class DataSet implements ArrayAccess, Countable {
    private array $rows = [];
    public function offsetExists(mixed $k): bool  { return isset($this->rows[$k]); }
    public function offsetGet(mixed $k): mixed    { return $this->rows[$k]; }
    public function offsetSet(mixed $k, mixed $v): void { $this->rows[] = $v; }
    public function offsetUnset(mixed $k): void   { unset($this->rows[$k]); }
    public function count(): int { return count($this->rows); }
}
$ds = new DataSet();
$ds[] = ['id' => 1]; $ds[] = ['id' => 2]; $ds[] = ['id' => 3];
echo count($ds);
"#,
    );
}

// ── JsonSerializable ──────────────────────────────────────────

#[test]
fn json_serializable_basic() {
    compile_ok(
        r#"<?php
class Money implements JsonSerializable {
    public function __construct(private int $cents, private string $currency = 'USD') {}
    public function jsonSerialize(): array {
        return ['amount' => $this->cents / 100, 'currency' => $this->currency];
    }
}
$m = new Money(1999, 'EUR');
echo json_encode($m);
"#,
    );
}

#[test]
fn json_serializable_nested() {
    compile_ok(
        r#"<?php
class Address implements JsonSerializable {
    public function __construct(public string $city, public string $country) {}
    public function jsonSerialize(): array { return ['city' => $this->city, 'country' => $this->country]; }
}
class Person implements JsonSerializable {
    public function __construct(public string $name, public Address $address) {}
    public function jsonSerialize(): array {
        return ['name' => $this->name, 'address' => $this->address];
    }
}
$p = new Person('Alice', new Address('Paris', 'FR'));
echo json_encode($p);
"#,
    );
}

// ── Stringable interface ──────────────────────────────────────

#[test]
fn stringable_basic() {
    compile_ok(
        r#"<?php
class Version implements Stringable {
    public function __construct(
        private int $major,
        private int $minor,
        private int $patch
    ) {}
    public function __toString(): string { return "{$this->major}.{$this->minor}.{$this->patch}"; }
}
function printVersion(Stringable $v): void { echo (string)$v; }
printVersion(new Version(1, 2, 3));
"#,
    );
}

// ── RecursiveIterator ─────────────────────────────────────────

#[test]
fn recursive_array_iterator() {
    compile_ok(
        r#"<?php
$tree = ['a', ['b', 'c'], ['d', ['e', 'f']]];
$it = new RecursiveIteratorIterator(
    new RecursiveArrayIterator($tree)
);
$leaves = [];
foreach ($it as $leaf) { $leaves[] = $leaf; }
echo implode(',', $leaves);
"#,
    );
}

#[test]
fn recursive_directory_iterator_stub() {
    compile_ok(
        r#"<?php
// RecursiveDirectoryIterator usage pattern
class TreeNode implements RecursiveIterator {
    private int $pos = 0;
    public function __construct(private array $children) {}
    public function current(): mixed  { return $this->children[$this->pos]; }
    public function key(): int        { return $this->pos; }
    public function next(): void      { $this->pos++; }
    public function rewind(): void    { $this->pos = 0; }
    public function valid(): bool     { return $this->pos < count($this->children); }
    public function hasChildren(): bool   { return is_array($this->current()); }
    public function getChildren(): static { return new static($this->current()); }
}
$tree = new TreeNode(['a', 'b', 'c']);
$items = [];
foreach ($tree as $item) { if (!is_array($item)) $items[] = $item; }
echo implode(',', $items);
"#,
    );
}

// ── AppendIterator ────────────────────────────────────────────

#[test]
fn append_iterator() {
    compile_ok(
        r#"<?php
$it1 = new ArrayIterator([1, 2, 3]);
$it2 = new ArrayIterator([4, 5, 6]);
$combined = new AppendIterator();
$combined->append($it1);
$combined->append($it2);
$result = [];
foreach ($combined as $v) { $result[] = $v; }
echo implode(',', $result);
"#,
    );
}

// ── CallbackFilterIterator ────────────────────────────────────

#[test]
fn callback_filter_iterator() {
    compile_ok(
        r#"<?php
$it = new ArrayIterator(range(1, 10));
$evens = new CallbackFilterIterator($it, fn($v) => $v % 2 === 0);
$result = [];
foreach ($evens as $v) { $result[] = $v; }
echo implode(',', $result);
"#,
    );
}

// ── LimitIterator ─────────────────────────────────────────────

#[test]
fn limit_iterator() {
    compile_ok(
        r#"<?php
$it = new ArrayIterator(range(0, 99));
$slice = new LimitIterator($it, 5, 5);
$result = [];
foreach ($slice as $v) { $result[] = $v; }
echo implode(',', $result);
"#,
    );
}

// ── iterator_to_array ─────────────────────────────────────────

#[test]
fn iterator_to_array_basic() {
    compile_ok(
        r#"<?php
$it = new ArrayIterator([3, 1, 4, 1, 5, 9]);
$arr = iterator_to_array($it, false);
sort($arr);
echo implode(',', $arr);
"#,
    );
}

#[test]
fn iterator_count() {
    compile_ok(
        r#"<?php
$it = new ArrayIterator([10, 20, 30, 40, 50]);
echo iterator_count($it);
"#,
    );
}

// ── Iterator protocol runtime (`php_cases!`) ────────────────────

crate::php_cases! {
    iterator_manual_next_advances_custom_class => {
        r#"<?php
class Three implements Iterator {
    private int $i = 0;
    public function current(): int { return $this->i; }
    public function key(): int { return $this->i; }
    public function next(): void { $this->i++; }
    public function rewind(): void { $this->i = 0; }
    public function valid(): bool { return $this->i < 3; }
}
$c = new Three();
$out = [];
while ($c->valid()) { $out[] = $c->current(); $c->next(); }
echo implode(',', $out);
"#,
        ["0,1,2"]
    };

    iterator_foreach_rewinds_before_loop => {
        r#"<?php
class Letters implements Iterator {
    private int $p = 0;
    private array $v = ['a', 'b'];
    public function current(): string { return $this->v[$this->p]; }
    public function key(): int { return $this->p; }
    public function next(): void { $this->p++; }
    public function rewind(): void { $this->p = 0; }
    public function valid(): bool { return $this->p < count($this->v); }
}
$it = new Letters();
foreach ($it as $ch) { echo $ch; }
"#,
        ["ab"]
    };

    arrayiterator_offset_access => {
        r#"<?php
$it = new ArrayIterator(['x' => 1, 'y' => 2]);
echo $it['x'] . $it['y'];
"#,
        ["12"]
    };

    iterator_to_array_preserves_values => {
        r#"<?php
$it = new ArrayIterator([3, 1, 2]);
echo implode(',', iterator_to_array($it));
"#,
        ["3,1,2"]
    };

    iterator_count_on_arrayiterator => {
        r#"<?php
echo iterator_count(new ArrayIterator([1, 2, 3]));
"#,
        ["3"]
    };

    emptyiterator_is_not_valid => {
        r#"<?php
$it = new EmptyIterator();
echo $it->valid() ? 'yes' : 'no';
"#,
        ["no"]
    };

    cachingiterator_has_next_after_first => {
        r#"<?php
$base = new ArrayIterator([1, 2]);
$cache = new CachingIterator($base);
$cache->next();
echo $cache->hasNext() ? 'more' : 'done';
"#,
        ["more"]
    };

    limititerator_caps_elements => {
        r#"<?php
$base = new ArrayIterator([1, 2, 3, 4, 5]);
$lim = new LimitIterator($base, 1, 2);
echo implode(',', iterator_to_array($lim));
"#,
        ["2,3"]
    };

    appenditerator_chains_two_iterators => {
        r#"<?php
$app = new AppendIterator();
$app->append(new ArrayIterator([1, 2]));
$app->append(new ArrayIterator([3]));
echo implode('', iterator_to_array($app));
"#,
        ["123"]
    };

    iteratoraggregate_yields_from_getiterator => {
        r#"<?php
class Bag implements IteratorAggregate {
    public function __construct(private array $items) {}
    public function getIterator(): Traversable {
        return new ArrayIterator($this->items);
    }
}
echo implode(',', iterator_to_array(new Bag([4, 5])));
"#,
        ["4,5"]
    };
}
