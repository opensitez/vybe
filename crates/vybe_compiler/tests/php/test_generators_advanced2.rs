use super::helpers::run_prints;

// ── yield from delegation ─────────────────────────────────────

#[test]
fn yield_from_array() {
    assert_eq!(
        run_prints(
            r#"<?php
function gen(): Generator { yield from [1, 2, 3]; }
echo implode(',', iterator_to_array(gen()));
"#
        ),
        vec!["1,2,3"]
    );
}
#[test]
fn yield_from_another_generator() {
    assert_eq!(
        run_prints(
            r#"<?php
function inner(): Generator { yield 'a'; yield 'b'; }
function outer(): Generator { yield from inner(); yield 'c'; }
echo implode('', iterator_to_array(outer()));
"#
        ),
        vec!["abc"]
    );
}
#[test]
fn yield_from_return_value() {
    assert_eq!(
        run_prints(
            r#"<?php
function child(): Generator { yield 1; return 'child_done'; }
function parent_gen(): Generator {
    $result = yield from child();
    yield $result;
}
echo implode(',', iterator_to_array(parent_gen()));
"#
        ),
        vec!["1,child_done"]
    );
}

// ── Generator::send() ─────────────────────────────────────────

#[test]
fn generator_send_bidirectional() {
    assert_eq!(
        run_prints(
            r#"<?php
function accumulator(): Generator {
    $total = 0;
    while (true) {
        $n = yield $total;
        if ($n === null) break;
        $total += $n;
    }
}
$g = accumulator();
$g->current();
echo $g->send(5) . ',' . $g->send(3) . ',' . $g->send(10);
"#
        ),
        vec!["5,8,18"]
    );
}
#[test]
fn generator_first_send_must_be_null() {
    assert_eq!(
        run_prints(
            r#"<?php
function gen(): Generator { $v = yield 'first'; yield "got:$v"; }
$g = gen();
echo $g->current() . ',';
echo $g->send('hello');
"#
        ),
        vec!["first,got:hello"]
    );
}

// ── Generator keys ────────────────────────────────────────────

#[test]
fn generator_key_value_pairs() {
    assert_eq!(
        run_prints(
            r#"<?php
function indexed(): Generator {
    yield 'a' => 1;
    yield 'b' => 2;
    yield 'c' => 3;
}
foreach (indexed() as $k => $v) echo $k . $v;
"#
        ),
        vec!["a1b2c3"]
    );
}
#[test]
fn generator_preserves_keys_false() {
    assert_eq!(
        run_prints(
            r#"<?php
function gen(): Generator { yield 'x' => 10; yield 'y' => 20; }
echo implode(',', iterator_to_array(gen(), false));
"#
        ),
        vec!["10,20"]
    );
}

// ── Generator return value ────────────────────────────────────

#[test]
fn generator_get_return() {
    assert_eq!(
        run_prints(
            r#"<?php
function gen(): Generator { yield 1; yield 2; return 'done'; }
$g = gen();
iterator_to_array($g);
echo $g->getReturn();
"#
        ),
        vec!["done"]
    );
}

// ── Generator exception handling ─────────────────────────────

#[test]
fn generator_throw_caught_inside() {
    assert_eq!(
        run_prints(
            r#"<?php
function resilient(): Generator {
    try {
        yield 1;
    } catch (RuntimeException $e) {
        yield 'caught:' . $e->getMessage();
    }
    yield 2;
}
$g = resilient();
echo $g->current() . ',';
$g->throw(new RuntimeException('boom'));
echo $g->current() . ',';
$g->next();
echo $g->current();
"#
        ),
        vec!["1,caught:boom,2"]
    );
}

// ── Infinite generators ───────────────────────────────────────

#[test]
fn fibonacci_generator() {
    assert_eq!(
        run_prints(
            r#"<?php
function fibonacci(): Generator {
    [$a, $b] = [0, 1];
    while (true) { yield $a; [$a, $b] = [$b, $a + $b]; }
}
$gen = fibonacci();
$result = [];
for ($i = 0; $i < 8; $i++) { $result[] = $gen->current(); $gen->next(); }
echo implode(',', $result);
"#
        ),
        vec!["0,1,1,2,3,5,8,13"]
    );
}
#[test]
fn generator_range_lazy() {
    assert_eq!(
        run_prints(
            r#"<?php
function lazyRange(int $start, int $end, int $step = 1): Generator {
    for ($i = $start; $i <= $end; $i += $step) yield $i;
}
$sum = 0;
foreach (lazyRange(1, 100) as $n) $sum += $n;
echo $sum;
"#
        ),
        vec!["5050"]
    );
}

// ── Generator in functional context ──────────────────────────

#[test]
fn generator_pipeline() {
    assert_eq!(
        run_prints(
            r#"<?php
function doubled(Generator $g): Generator { foreach ($g as $v) yield $v * 2; }
function filtered(Generator $g, callable $fn): Generator { foreach ($g as $v) if ($fn($v)) yield $v; }
function gen(): Generator { for ($i=1;$i<=5;$i++) yield $i; }
$pipeline = filtered(doubled(gen()), fn($n) => $n > 4);
echo implode(',', iterator_to_array($pipeline));
"#
        ),
        vec!["6,8,10"]
    );
}
#[test]
fn generator_as_coroutine_state_machine() {
    assert_eq!(
        run_prints(
            r#"<?php
function stateMachine(): Generator {
    $state = 'idle';
    while (true) {
        $cmd = yield $state;
        $state = match($cmd) {
            'start' => 'running',
            'pause' => 'paused',
            'stop'  => 'stopped',
            default => $state,
        };
        if ($state === 'stopped') return;
    }
}
$sm = stateMachine();
echo $sm->current() . ',';
echo $sm->send('start') . ',';
echo $sm->send('pause') . ',';
echo $sm->send('stop');
"#
        ),
        vec!["idle,running,paused,stopped"]
    );
}
