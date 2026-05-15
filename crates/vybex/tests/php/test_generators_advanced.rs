use super::helpers::run_prints;

// ── Generator send() ─────────────────────────────────────────────
#[test]
fn generator_send_basic() {
    assert_eq!(run_prints(r#"<?php
function accumulator() {
    $total = 0;
    while (true) {
        $value = yield $total;
        if ($value === null) break;
        $total += $value;
    }
}
$gen = accumulator();
$gen->current();  // start
$gen->send(10);
$gen->send(20);
echo $gen->send(30);
"#), &["60"]);
}

#[test]
fn generator_send_echo_each() {
    assert_eq!(run_prints(r#"<?php
function logger() {
    while (true) {
        $msg = yield;
        if ($msg === null) break;
        echo strtoupper($msg);
    }
}
$log = logger();
$log->current();
$log->send("hello");
$log->send("world");
"#), &["HELLO", "WORLD"]);
}

// ── Generator getReturn() ────────────────────────────────────────
#[test]
fn generator_return_value() {
    assert_eq!(run_prints(r#"<?php
function sumGenerator(array $numbers) {
    $sum = 0;
    foreach ($numbers as $n) {
        $sum += $n;
        yield $n;
    }
    return $sum;
}
$gen = sumGenerator([1, 2, 3, 4, 5]);
foreach ($gen as $val) {
    // consume all yielded values
}
echo $gen->getReturn();
"#), &["15"]);
}

// ── yield from delegation ────────────────────────────────────────
#[test]
fn yield_from_array() {
    assert_eq!(run_prints(r#"<?php
function inner() {
    yield 1;
    yield 2;
    yield 3;
}
function outer() {
    yield 0;
    yield from inner();
    yield 4;
}
$result = [];
foreach (outer() as $v) {
    $result[] = $v;
}
echo implode(",", $result);
"#), &["0,1,2,3,4"]);
}

#[test]
fn yield_from_nested() {
    assert_eq!(run_prints(r#"<?php
function leaves(array $tree) {
    foreach ($tree as $item) {
        if (is_array($item)) {
            yield from leaves($item);
        } else {
            yield $item;
        }
    }
}
$tree = [1, [2, 3], [4, [5, 6]]];
$flat = [];
foreach (leaves($tree) as $leaf) {
    $flat[] = $leaf;
}
echo implode(",", $flat);
"#), &["1,2,3,4,5,6"]);
}

// ── Generator pipelines ──────────────────────────────────────────
#[test]
fn generator_pipeline_map_filter() {
    assert_eq!(run_prints(r#"<?php
function rangeGen(int $start, int $end) {
    for ($i = $start; $i <= $end; $i++) {
        yield $i;
    }
}
function filterGen($gen, callable $pred) {
    foreach ($gen as $val) {
        if ($pred($val)) yield $val;
    }
}
function mapGen($gen, callable $fn) {
    foreach ($gen as $val) {
        yield $fn($val);
    }
}
$numbers = rangeGen(1, 10);
$evens = filterGen($numbers, fn($n) => $n % 2 == 0);
$doubled = mapGen($evens, fn($n) => $n * 2);
$result = [];
foreach ($doubled as $v) {
    $result[] = $v;
}
echo implode(",", $result);
"#), &["4,8,12,16,20"]);
}

// ── Generator with keys ──────────────────────────────────────────
#[test]
fn generator_key_value_pairs() {
    assert_eq!(run_prints(r#"<?php
function csvRows(string $csv) {
    $lines = explode("\n", trim($csv));
    $headers = str_getcsv(array_shift($lines));
    foreach ($lines as $i => $line) {
        $values = str_getcsv($line);
        yield $i => array_combine($headers, $values);
    }
}
$csv = "name,age\nAlice,30\nBob,25";
foreach (csvRows($csv) as $idx => $row) {
    echo "$idx: {$row['name']} is {$row['age']}";
}
"#), &["0: Alice is 30", "1: Bob is 25"]);
}

#[test]
fn generator_yields_map_values() {
    assert_eq!(run_prints(r#"<?php
function csvRows(string $csv) {
    $lines = explode("\n", trim($csv));
    $headers = str_getcsv(array_shift($lines));
    foreach ($lines as $line) {
        $values = str_getcsv($line);
        yield array_combine($headers, $values);
    }
}
$csv = "name,age\nAlice,30\nBob,25";
foreach (csvRows($csv) as $row) {
    echo "{$row['name']} is {$row['age']}";
}
"#), &["Alice is 30", "Bob is 25"]);
}

// ── Generator valid/rewind ───────────────────────────────────────
#[test]
fn generator_valid_check() {
    assert_eq!(run_prints(r#"<?php
function countdown(int $from) {
    for ($i = $from; $i > 0; $i--) {
        yield $i;
    }
}
$gen = countdown(3);
echo $gen->valid() ? "valid" : "done";
$gen->next();
$gen->next();
$gen->next();
echo $gen->valid() ? "valid" : "done";
"#), &["valid", "done"]);
}

// ── Generator as lazy infinite sequence ──────────────────────────
#[test]
fn generator_fibonacci_lazy() {
    assert_eq!(run_prints(r#"<?php
function fibonacci() {
    $a = 0;
    $b = 1;
    while (true) {
        yield $a;
        [$a, $b] = [$b, $a + $b];
    }
}
$fib = fibonacci();
$result = [];
for ($i = 0; $i < 8; $i++) {
    $result[] = $fib->current();
    $fib->next();
}
echo implode(",", $result);
"#), &["0,1,1,2,3,5,8,13"]);
}

// ── Generator take helper ────────────────────────────────────────
#[test]
fn generator_take_n() {
    assert_eq!(run_prints(r#"<?php
function naturals() {
    $n = 1;
    while (true) {
        yield $n++;
    }
}
function take(Generator $gen, int $n): array {
    $result = [];
    for ($i = 0; $i < $n; $i++) {
        $result[] = $gen->current();
        $gen->next();
    }
    return $result;
}
echo implode(",", take(naturals(), 5));
"#), &["1,2,3,4,5"]);
}

// ── yield from return value ──────────────────────────────────────
#[test]
fn yield_from_captures_return() {
    assert_eq!(run_prints(r#"<?php
function inner() {
    yield "a";
    yield "b";
    return "done";
}
function outer() {
    $result = yield from inner();
    echo $result;
}
foreach (outer() as $v) {
    echo $v;
}
"#), &["a", "b", "done"]);
}

// ── Generator with exception handling ────────────────────────────
#[test]
fn generator_throw() {
    assert_eq!(run_prints(r#"<?php
function safeGenerator() {
    try {
        yield 1;
        yield 2;
        yield 3;
    } catch (Exception $e) {
        echo "caught: " . $e->getMessage();
    }
}
$gen = safeGenerator();
echo $gen->current();
$gen->next();
echo $gen->current();
$gen->throw(new Exception("stop"));
"#), &["1", "2", "caught: stop"]);
}

#[test]
fn generator_throw_before_first_yield() {
    assert_eq!(run_prints(r#"<?php
function handled() {
    try {
        yield "ready";
    } catch (Exception $e) {
        echo "caught";
        yield "handled";
    }
}
$gen = handled();
echo $gen->throw(new Exception("stop"));
"#), &["caught", "handled"]);
}

#[test]
fn generator_variadic_throw_before_first_yield() {
    assert_eq!(run_prints(r#"<?php
function handled($head, ...$rest) {
    try {
        yield count($rest);
    } catch (Exception $e) {
        echo implode(',', $rest);
        yield $e->getMessage();
    }
}
$gen = handled('a', 'b', 'c');
echo $gen->throw(new Exception('stop'));
"#), &["b,c", "stop"]);
}

// ── Coroutine pattern ────────────────────────────────────────────
#[test]
fn coroutine_echo_back() {
    assert_eq!(run_prints(r#"<?php
function echoBack() {
    $received = [];
    while (true) {
        $input = yield;
        if ($input === "done") break;
        $received[] = $input;
    }
    return implode(",", $received);
}
$co = echoBack();
$co->current();
$co->send("hello");
$co->send("world");
$co->send("done");
echo $co->getReturn();
"#), &["hello,world"]);
}

// ── Generator with finally ───────────────────────────────────────
#[test]
fn generator_finally_cleanup() {
    assert_eq!(run_prints(r#"<?php
function resource() {
    echo "open";
    try {
        yield 1;
        yield 2;
    } finally {
        echo "close";
    }
}
$gen = resource();
echo $gen->current();
$gen->next();
echo $gen->current();
$gen = null; // drop triggers finally
"#), &["open", "1", "2", "close"]);
}
