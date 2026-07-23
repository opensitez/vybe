use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP SPL: InfiniteIterator Cyclic Traversal & Outer Iterator
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_spl_infinite_iterator_cycles_repeatedly() {
    let out = run_prints(
        r##"<?php
$arr = new ArrayIterator(["red", "blue"]);
$inf = new InfiniteIterator($arr);

$out = [];
$count = 0;
foreach ($inf as $val) {
    $out[] = $val;
    $count++;
    if ($count >= 5) break;
}
echo implode(",", $out);
"##,
    );
    assert_eq!(out, vec!["red,blue,red,blue,red"]);
}

#[test]
fn test_php_spl_infinite_iterator_get_inner_iterator() {
    let out = run_prints(
        r##"<?php
$inner = new ArrayIterator(["a", "b"]);
$inf = new InfiniteIterator($inner);

echo $inf->getInnerIterator() === $inner ? "INNER_MATCH" : "FAIL";
"##,
    );
    assert_eq!(out, vec!["INNER_MATCH"]);
}

#[test]
fn test_php_spl_infinite_iterator_limit_iterator_combination() {
    let out = run_prints(
        r##"<?php
$arr = new ArrayIterator([1, 2, 3]);
$inf = new InfiniteIterator($arr);
$limited = new LimitIterator($inf, 0, 7);

$vals = iterator_to_array($limited, false);
echo implode("-", $vals);
"##,
    );
    assert_eq!(out, vec!["1-2-3-1-2-3-1"]);
}

#[test]
fn test_php_spl_infinite_iterator_next_wraps_around() {
    compile_ok(
        r##"<?php
$arr = new ArrayIterator(["x"]);
$inf = new InfiniteIterator($arr);
$inf->rewind();
$inf->next();
echo $inf->valid() && $inf->current() === "x" ? "WRAP_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_infinite_iterator_empty_inner_iterator_halts() {
    compile_ok(
        r##"<?php
$arr = new ArrayIterator([]);
$inf = new InfiniteIterator($arr);
$inf->rewind();
echo !$inf->valid() ? "EMPTY_INFINITE_HALTS" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_infinite_iterator_key_repetition() {
    compile_ok(
        r##"<?php
$arr = new ArrayIterator(["k1" => "v1", "k2" => "v2"]);
$inf = new InfiniteIterator($arr);
$keys = [];
$i = 0;
foreach ($inf as $k => $v) {
    $keys[] = $k;
    if (++$i >= 4) break;
}
echo implode(",", $keys) === "k1,k2,k1,k2" ? "KEYS_REPEAT_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_infinite_iterator_custom_iterator_aggregate() {
    compile_ok(
        r##"<?php
class SimpleCollection implements IteratorAggregate {
    public function getIterator(): Traversable {
        return new ArrayIterator([10, 20]);
    }
}
$coll = new SimpleCollection();
$inf = new InfiniteIterator($coll->getIterator());
$inf->rewind();
echo $inf->current() === 10 ? "AGGREGATE_INFINITE_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_infinite_iterator_multiple_rewinds() {
    compile_ok(
        r##"<?php
$arr = new ArrayIterator(["a", "b"]);
$inf = new InfiniteIterator($arr);
$inf->next();
$inf->rewind();
echo $inf->current() === "a" ? "MULTIPLE_REWIND_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_infinite_iterator_instanceof_iterator() {
    compile_ok(
        r##"<?php
$arr = new ArrayIterator([1]);
$inf = new InfiniteIterator($arr);
echo ($inf instanceof Iterator) ? "INSTANCEOF_ITERATOR" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_infinite_iterator_current_after_several_steps() {
    compile_ok(
        r##"<?php
$arr = new ArrayIterator(["one", "two"]);
$inf = new InfiniteIterator($arr);
$inf->rewind();
$inf->next(); // two
$inf->next(); // one
$inf->next(); // two
echo $inf->current() === "two" ? "STEP_POSITION_OK" : "FAIL";
"##,
    );
}
