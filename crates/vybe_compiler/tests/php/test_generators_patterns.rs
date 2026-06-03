use super::helpers::run_prints;

// ── Generator basics recap (non-duplicate angles) ─────────────

#[test]
fn generator_early_return() {
    assert_eq!(
        run_prints(
            r#"<?php
function limited(): Generator {
    yield 1; yield 2;
    if (true) return;
    yield 3;
}
echo implode(',', iterator_to_array(limited()));
"#
        ),
        vec!["1,2"]
    );
}
#[test]
fn generator_valid_after_all_consumed() {
    assert_eq!(
        run_prints(
            r#"<?php
function gen(): Generator { yield 1; yield 2; }
$g = gen();
iterator_to_array($g);
echo $g->valid() ? 'valid' : 'done';
"#
        ),
        vec!["done"]
    );
}
#[test]
fn generator_count_not_applicable() {
    assert_eq!(
        run_prints(
            r#"<?php
function gen(): Generator { yield 1; yield 2; yield 3; }
$arr = iterator_to_array(gen());
echo count($arr);
"#
        ),
        vec!["3"]
    );
}

// ── Generator as lazy pipeline ────────────────────────────────

#[test]
fn generator_take_n() {
    assert_eq!(
        run_prints(
            r#"<?php
function naturals(): Generator { $n = 1; while (true) yield $n++; }
function take(Generator $g, int $n): array {
    $result = [];
    for ($i = 0; $i < $n; $i++) { $result[] = $g->current(); $g->next(); }
    return $result;
}
echo implode(',', take(naturals(), 5));
"#
        ),
        vec!["1,2,3,4,5"]
    );
}
#[test]
fn generator_map_lazy() {
    assert_eq!(
        run_prints(
            r#"<?php
function mapGen(callable $fn, Generator $g): Generator { foreach ($g as $v) yield $fn($v); }
function genRange(int $a, int $b): Generator { for ($i=$a;$i<=$b;$i++) yield $i; }
$doubled = mapGen(fn($n) => $n*2, genRange(1, 5));
echo implode(',', iterator_to_array($doubled));
"#
        ),
        vec!["2,4,6,8,10"]
    );
}
#[test]
fn generator_filter_lazy() {
    assert_eq!(
        run_prints(
            r#"<?php
function filterGen(callable $fn, Generator $g): Generator { foreach ($g as $v) if ($fn($v)) yield $v; }
function genNums(): Generator { yield from range(1, 10); }
$odds = filterGen(fn($n) => $n % 2 !== 0, genNums());
echo implode(',', iterator_to_array($odds));
"#
        ),
        vec!["1,3,5,7,9"]
    );
}

// ── yield from with return value ──────────────────────────────

#[test]
fn yield_from_nested_return() {
    assert_eq!(
        run_prints(
            r#"<?php
function inner2(): Generator { yield 1; yield 2; return 'inner_done'; }
function outer2(): Generator {
    $result = yield from inner2();
    echo "inner returned: $result\n";
    yield 3;
}
$g = outer2();
iterator_to_array($g);
"#
        ),
        vec!["inner returned: inner_done\n"]
    );
}

// ── Generator with exceptions ─────────────────────────────────

#[test]
fn generator_finally_on_early_close() {
    assert_eq!(
        run_prints(
            r#"<?php
function withCleanup(): Generator {
    try {
        yield 1;
        yield 2;
        yield 3;
    } finally {
        echo 'cleanup';
    }
}
$g = withCleanup();
echo $g->current() . ',';
$g->next();
$g = null;
"#
        ),
        vec!["1,cleanup"]
    );
}

// ── Generator rewind ─────────────────────────────────────────

#[test]
fn generator_rewind_restarts() {
    assert_eq!(
        run_prints(
            r#"<?php
function abc(): Generator { yield 'a'; yield 'b'; yield 'c'; }
$g = abc();
echo $g->current();
$g->rewind();
echo $g->current();
"#
        ),
        vec!["aa"]
    );
}

// ── Generator with complex types ─────────────────────────────

#[test]
fn generator_yields_objects() {
    assert_eq!(
        run_prints(
            r#"<?php
class Item { public function __construct(public string $name) {} }
function items(): Generator { yield new Item('foo'); yield new Item('bar'); }
$names = [];
foreach (items() as $item) $names[] = $item->name;
echo implode(',', $names);
"#
        ),
        vec!["foo,bar"]
    );
}
#[test]
fn generator_accumulates_state() {
    assert_eq!(
        run_prints(
            r#"<?php
function runningTotal(array $nums): Generator {
    $total = 0;
    foreach ($nums as $n) { $total += $n; yield $total; }
}
echo implode(',', iterator_to_array(runningTotal([1,2,3,4,5])));
"#
        ),
        vec!["1,3,6,10,15"]
    );
}
