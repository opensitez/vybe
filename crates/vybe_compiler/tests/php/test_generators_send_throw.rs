use super::helpers::{compile_ok, run_prints};

// ── Generator::send() — bidirectional communication ───────────

#[test]
fn generator_send_receives_value() {
    assert_eq!(
        run_prints(
            r#"<?php
function accumulator(): Generator {
    $total = 0;
    while (true) {
        $val = yield $total;
        if ($val === null) break;
        $total += $val;
    }
}
$gen = accumulator();
$gen->current();
$gen->send(10);
$gen->send(20);
echo $gen->send(5);
"#
        ),
        vec!["35"]
    );
}

#[test]
fn generator_send_initial_null() {
    assert_eq!(
        run_prints(
            r#"<?php
function echoSent(): Generator {
    $val = yield 'ready';
    yield $val;
}
$gen = echoSent();
$gen->current();
echo $gen->send('hello');
"#
        ),
        vec!["hello"]
    );
}

#[test]
fn generator_send_controls_loop() {
    assert_eq!(
        run_prints(
            r#"<?php
function counter(): Generator {
    $i = 0;
    while (true) {
        $step = yield $i;
        $i += $step ?? 1;
    }
}
$g = counter();
$g->current();
$g->send(2);
$g->send(3);
echo $g->send(5);
"#
        ),
        vec!["10"]
    );
}

// ── Generator::getReturn() ────────────────────────────────────

#[test]
fn generator_return_value_accessible_via_get_return() {
    assert_eq!(
        run_prints(
            r#"<?php
function sumToN(int $n): Generator {
    $sum = 0;
    for ($i = 1; $i <= $n; $i++) {
        $sum += $i;
        yield $i;
    }
    return $sum;
}
$gen = sumToN(4);
foreach ($gen as $_) {}
echo $gen->getReturn();
"#
        ),
        vec!["10"]
    );
}

#[test]
fn generator_return_value_is_null_if_no_return() {
    assert_eq!(
        run_prints(
            r#"<?php
function simpleYield(): Generator {
    yield 1;
    yield 2;
}
$gen = simpleYield();
foreach ($gen as $_) {}
echo var_export($gen->getReturn(), true);
"#
        ),
        vec!["NULL"]
    );
}

#[test]
fn generator_return_string_value() {
    assert_eq!(
        run_prints(
            r#"<?php
function produce(): Generator {
    yield 'item';
    return 'done';
}
$gen = produce();
$gen->current();
$gen->next();
echo $gen->getReturn();
"#
        ),
        vec!["done"]
    );
}

// ── Generator::throw() ────────────────────────────────────────

#[test]
fn generator_throw_triggers_exception_inside() {
    assert_eq!(
        run_prints(
            r#"<?php
function resilient(): Generator {
    try {
        yield 'before';
    } catch (RuntimeException $e) {
        yield 'caught: ' . $e->getMessage();
    }
}
$gen = resilient();
$gen->current();
echo $gen->throw(new RuntimeException('oops'));
"#
        ),
        vec!["caught: oops"]
    );
}

#[test]
fn generator_throw_propagates_if_uncaught() {
    assert_eq!(
        run_prints(
            r#"<?php
function fragile(): Generator {
    yield 'start';
}
$gen = fragile();
$gen->current();
try {
    $gen->throw(new Exception('boom'));
} catch (Exception $e) {
    echo $e->getMessage();
}
"#
        ),
        vec!["boom"]
    );
}

// ── Generator with finally ────────────────────────────────────

#[test]
fn generator_finally_runs_on_early_close() {
    assert_eq!(
        run_prints(
            r#"<?php
function withCleanup(): Generator {
    try {
        yield 1;
        yield 2;
    } finally {
        echo "cleanup";
    }
}
$gen = withCleanup();
$gen->current();
$gen = null;
"#
        ),
        vec!["cleanup"]
    );
}

#[test]
fn generator_finally_runs_on_normal_completion() {
    assert_eq!(
        run_prints(
            r#"<?php
function withFinally(): Generator {
    try {
        yield 'a';
    } finally {
        echo "done";
    }
}
$gen = withFinally();
foreach ($gen as $_) {}
"#
        ),
        vec!["done"]
    );
}

// ── yield from delegation ─────────────────────────────────────

#[test]
fn yield_from_array_delegation() {
    assert_eq!(
        run_prints(
            r#"<?php
function fromArray(): Generator {
    yield from [1, 2, 3];
}
echo implode(',', iterator_to_array(fromArray()));
"#
        ),
        vec!["1,2,3"]
    );
}

#[test]
fn yield_from_generator_delegation() {
    assert_eq!(
        run_prints(
            r#"<?php
function inner(): Generator { yield 'a'; yield 'b'; }
function outer(): Generator {
    yield 'start';
    yield from inner();
    yield 'end';
}
echo implode(',', iterator_to_array(outer()));
"#
        ),
        vec!["start,a,b,end"]
    );
}

#[test]
fn yield_from_returns_inner_generator_return() {
    assert_eq!(
        run_prints(
            r#"<?php
function inner(): Generator {
    yield 1;
    return 'inner_done';
}
function outer(): Generator {
    $result = yield from inner();
    yield $result;
}
$gen = outer();
$gen->current();
$gen->next();
echo $gen->current();
"#
        ),
        vec!["inner_done"]
    );
}

// ── Generator key => value ────────────────────────────────────

#[test]
fn generator_yield_key_value_pairs() {
    assert_eq!(
        run_prints(
            r#"<?php
function kvPairs(): Generator {
    yield 'a' => 1;
    yield 'b' => 2;
    yield 'c' => 3;
}
$result = [];
foreach (kvPairs() as $k => $v) {
    $result[] = "$k=$v";
}
echo implode(',', $result);
"#
        ),
        vec!["a=1,b=2,c=3"]
    );
}

// ── Generator as lazy range ───────────────────────────────────

#[test]
fn generator_lazy_range_with_step() {
    assert_eq!(
        run_prints(
            r#"<?php
function lazyRange(int $start, int $end, int $step = 1): Generator {
    for ($i = $start; $i <= $end; $i += $step) {
        yield $i;
    }
}
echo implode(',', iterator_to_array(lazyRange(0, 10, 3)));
"#
        ),
        vec!["0,3,6,9"]
    );
}

// ── Generator valid/current/next state machine ────────────────

#[test]
fn generator_valid_returns_false_after_exhaustion() {
    assert_eq!(
        run_prints(
            r#"<?php
function two(): Generator { yield 1; yield 2; }
$gen = two();
$gen->current();
$gen->next();
$gen->next();
echo $gen->valid() ? 'yes' : 'no';
"#
        ),
        vec!["no"]
    );
}

#[test]
fn generator_rewind_does_not_reset() {
    assert_eq!(
        run_prints(
            r#"<?php
function seq(): Generator { yield 1; yield 2; }
$gen = seq();
$gen->next();
$gen->rewind();
echo $gen->current();
"#
        ),
        vec!["1"]
    );
}

// ── Infinite generator with early break ──────────────────────

#[test]
fn infinite_generator_take_n_items() {
    assert_eq!(
        run_prints(
            r#"<?php
function naturals(): Generator {
    $n = 1;
    while (true) { yield $n++; }
}
$result = [];
foreach (naturals() as $v) {
    $result[] = $v;
    if (count($result) >= 5) break;
}
echo implode(',', $result);
"#
        ),
        vec!["1,2,3,4,5"]
    );
}

// ── Generator with complex send protocol ─────────────────────

#[test]
fn generator_calculator_protocol() {
    assert_eq!(
        run_prints(
            r#"<?php
function calculator(): Generator {
    $result = 0;
    while (true) {
        [$op, $val] = yield $result;
        $result = match($op) {
            '+' => $result + $val,
            '-' => $result - $val,
            '*' => $result * $val,
            default => $result,
        };
    }
}
$calc = calculator();
$calc->current();
$calc->send(['+', 10]);
$calc->send(['*', 3]);
echo $calc->send(['-', 5]);
"#
        ),
        vec!["25"]
    );
}

// ── iterator_to_array behavior ────────────────────────────────

#[test]
fn iterator_to_array_preserve_keys_true() {
    assert_eq!(
        run_prints(
            r#"<?php
function pairs(): Generator {
    yield 'x' => 10;
    yield 'y' => 20;
}
$arr = iterator_to_array(pairs(), true);
echo $arr['x'] . ',' . $arr['y'];
"#
        ),
        vec!["10,20"]
    );
}

#[test]
fn iterator_to_array_preserve_keys_false_reindexes() {
    assert_eq!(
        run_prints(
            r#"<?php
function repeated(): Generator {
    yield 'a' => 1;
    yield 'a' => 2;
}
$arr = iterator_to_array(repeated(), false);
echo count($arr) . ',' . $arr[0] . ',' . $arr[1];
"#
        ),
        vec!["2,1,2"]
    );
}

// ── Generator exception handling chain ───────────────────────

#[test]
fn generator_catch_rethrow_pattern() {
    assert_eq!(
        run_prints(
            r#"<?php
function safeGen(): Generator {
    try {
        yield 'step1';
        yield 'step2';
    } catch (InvalidArgumentException $e) {
        yield 'handled: ' . $e->getMessage();
    }
}
$gen = safeGen();
$gen->current();
echo $gen->throw(new InvalidArgumentException('bad input'));
"#
        ),
        vec!["handled: bad input"]
    );
}

#[test]
fn generator_multiple_yields_after_send() {
    assert_eq!(
        run_prints(
            r#"<?php
function multiStep(): Generator {
    $a = yield 'first';
    $b = yield 'second';
    yield "$a+$b=" . ($a + $b);
}
$gen = multiStep();
$gen->current();
$gen->send(3);
echo $gen->send(7);
"#
        ),
        vec!["3+7=10"]
    );
}
