use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Generators — yield, yield from, $gen->send(), $gen->throw(), $gen->getReturn()
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_generator_basic_yield_sequence() {
    let out = run_prints(
        r#"<?php
function rangeGen($start, $end) {
    for ($i = $start; $i <= $end; $i++) {
        yield $i;
    }
}

$items = [];
foreach (rangeGen(1, 4) as $num) {
    $items[] = $num;
}
echo implode("-", $items);
"#,
    );
    assert_eq!(out, vec!["1-2-3-4"]);
}

#[test]
fn test_php_generator_yield_key_value_pairs() {
    let out = run_prints(
        r#"<?php
function keyValueGen() {
    yield "a" => 10;
    yield "b" => 20;
}

foreach (keyValueGen() as $k => $v) {
    echo "$k=$v ";
}
"#,
    );
    assert_eq!(out, vec!["a=10 b=20 "]);
}

#[test]
fn test_php_generator_yield_from_delegation() {
    let out = run_prints(
        r#"<?php
function innerGen() {
    yield 2;
    yield 3;
}

function outerGen() {
    yield 1;
    yield from innerGen();
    yield 4;
}

echo implode(",", iterator_to_array(outerGen()));
"#,
    );
    assert_eq!(out, vec!["1,2,3,4"]);
}

#[test]
fn test_php_generator_send_bidirectional_communication() {
    let out = run_prints(
        r#"<?php
function loggerGen() {
    $val = yield "ready";
    yield "received: $val";
}

$g = loggerGen();
echo $g->current() . " | ";
echo $g->send("hello_generator");
"#,
    );
    assert_eq!(out, vec!["ready | received: hello_generator"]);
}

#[test]
fn test_php_generator_return_value_get_return() {
    let out = run_prints(
        r#"<?php
function countAndReturn() {
    yield 10;
    yield 20;
    return "done_counting";
}

$g = countAndReturn();
foreach ($g as $v) {}
echo $g->getReturn();
"#,
    );
    assert_eq!(out, vec!["done_counting"]);
}

#[test]
fn test_php_generator_throw_exception_into_generator() {
    compile_ok(
        r#"<?php
function exceptionGen() {
    try {
        yield "start";
    } catch (RuntimeException $e) {
        yield "handled: " . $e->getMessage();
    }
}

$g = exceptionGen();
echo $g->current();
echo $g->throw(new RuntimeException("Injected Error"));
"#,
    );
}

#[test]
fn test_php_generator_by_reference_yield() {
    compile_ok(
        r#"<?php
function &refGen(&$val) {
    yield $val;
}

$num = 100;
$g = refGen($num);
foreach ($g as &$v) {
    $v += 50;
}
echo $num;
"#,
    );
}

#[test]
fn test_php_yield_from_array_and_traversable() {
    compile_ok(
        r#"<?php
function delegateArray() {
    yield from [10, 20, 30];
    yield from new ArrayIterator([40, 50]);
}

$all = iterator_to_array(delegateArray(), false);
echo implode("+", $all);
"#,
    );
}

#[test]
fn test_php_generator_rewind_behavior() {
    compile_ok(
        r#"<?php
function gen() {
    yield 1;
    yield 2;
}

$g = gen();
$g->rewind();
echo $g->current();
$g->next();
echo $g->current();
"#,
    );
}

#[test]
fn test_php_generator_closed_state_inspection() {
    compile_ok(
        r#"<?php
function simpleGen() {
    yield 1;
}

$g = simpleGen();
foreach ($g as $v) {}
echo $g->valid() ? "VALID" : "EXHAUSTED";
"#,
    );
}
