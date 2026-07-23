use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Iterators & SPL — Iterator, IteratorAggregate, ArrayIterator, LimitIterator, CallbackFilterIterator, RecursiveIteratorIterator
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_custom_iterator_interface_implementation() {
    let out = run_prints(
        r#"<?php
class NumberRange implements Iterator {
    private int $position = 0;
    public function __construct(private int $start, private int $end) {
        $this->position = $start;
    }
    public function current(): mixed { return $this->position; }
    public function key(): mixed { return $this->position - $this->start; }
    public function next(): void { $this->position++; }
    public function rewind(): void { $this->position = $this->start; }
    public function valid(): bool { return $this->position <= $this->end; }
}

$range = new NumberRange(10, 12);
$out = [];
foreach ($range as $k => $v) {
    $out[] = "$k:$v";
}
echo implode(", ", $out);
"#,
    );
    assert_eq!(out, vec!["0:10, 1:11, 2:12"]);
}

#[test]
fn test_php_iterator_aggregate_get_iterator() {
    let out = run_prints(
        r#"<?php
class Collection implements IteratorAggregate {
    private array $items = ["a", "b", "c"];
    public function getIterator(): Traversable {
        return new ArrayIterator($this->items);
    }
}

$c = new Collection();
echo implode("-", iterator_to_array($c));
"#,
    );
    assert_eq!(out, vec!["a-b-c"]);
}

#[test]
fn test_php_callback_filter_iterator() {
    let out = run_prints(
        r#"<?php
$array = [1, 2, 3, 4, 5, 6];
$iterator = new ArrayIterator($array);
$filtered = new CallbackFilterIterator($iterator, fn($val) => $val % 2 === 0);

echo implode(",", iterator_to_array($filtered));
"#,
    );
    assert_eq!(out, vec!["2,4,6"]);
}

#[test]
fn test_php_limit_iterator_offset_and_count() {
    let out = run_prints(
        r#"<?php
$array = [10, 20, 30, 40, 50];
$it = new LimitIterator(new ArrayIterator($array), 1, 3); // offset 1, count 3
echo implode(",", iterator_to_array($it, false));
"#,
    );
    assert_eq!(out, vec!["20,30,40"]);
}

#[test]
fn test_php_array_iterator_ksort_and_asort() {
    compile_ok(
        r#"<?php
$it = new ArrayIterator(["b" => 2, "a" => 1, "c" => 3]);
$it->ksort();
print_r(iterator_to_array($it));
"#,
    );
}

#[test]
fn test_php_recursive_array_iterator_flattening() {
    compile_ok(
        r#"<?php
$data = [1, [2, 3], [4, [5, 6]]];
$it = new RecursiveIteratorIterator(new RecursiveArrayIterator($data));
$flat = [];
foreach ($it as $v) {
    $flat[] = $v;
}
echo implode(",", $flat);
"#,
    );
}

#[test]
fn test_php_infinite_iterator_looping() {
    compile_ok(
        r#"<?php
$it = new InfiniteIterator(new ArrayIterator([1, 2]));
$limit = new LimitIterator($it, 0, 5);
echo implode(",", iterator_to_array($limit, false));
"#,
    );
}

#[test]
fn test_php_multiple_iterator_synchronization() {
    compile_ok(
        r#"<?php
$mit = new MultipleIterator(MultipleIterator::MIT_NEED_ALL);
$mit->attachIterator(new ArrayIterator([1, 2]));
$mit->attachIterator(new ArrayIterator(["a", "b"]));

foreach ($mit as $pair) {
    echo $pair[0] . "-" . $pair[1] . "\n";
}
"#,
    );
}

#[test]
fn test_php_append_iterator_chaining() {
    compile_ok(
        r#"<?php
$app = new AppendIterator();
$app->append(new ArrayIterator([1, 2]));
$app->append(new ArrayIterator([3, 4]));
echo implode(",", iterator_to_array($app, false));
"#,
    );
}

#[test]
fn test_php_caching_iterator_lookahead() {
    compile_ok(
        r#"<?php
$cit = new CachingIterator(new ArrayIterator([1, 2, 3]), CachingIterator::FULL_CACHE);
foreach ($cit as $val) {
    if (!$cit->hasNext()) {
        echo "LAST:$val";
    }
}
"#,
    );
}
