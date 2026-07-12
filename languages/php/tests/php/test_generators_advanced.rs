use super::helpers::{compile_ok, run_prints};

// ── Generator send() ─────────────────────────────────────────────
#[test]
fn generator_send_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
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
"#
        ),
        &["60"]
    );
}

#[test]
fn generator_send_echo_each() {
    assert_eq!(
        run_prints(
            r#"<?php
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
"#
        ),
        &["HELLOWORLD"]
    );
}

// ── Generator getReturn() ────────────────────────────────────────
#[test]
fn generator_return_value() {
    assert_eq!(
        run_prints(
            r#"<?php
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
"#
        ),
        &["15"]
    );
}

// ── yield from delegation ────────────────────────────────────────
#[test]
fn yield_from_array() {
    assert_eq!(
        run_prints(
            r#"<?php
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
"#
        ),
        &["0,1,2,3,4"]
    );
}

#[test]
fn yield_from_nested() {
    assert_eq!(
        run_prints(
            r#"<?php
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
"#
        ),
        &["1,2,3,4,5,6"]
    );
}

// ── Generator pipelines ──────────────────────────────────────────
#[test]
fn generator_pipeline_map_filter() {
    assert_eq!(
        run_prints(
            r#"<?php
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
"#
        ),
        &["4,8,12,16,20"]
    );
}

// ── Generator with keys ──────────────────────────────────────────
#[test]
fn generator_key_value_pairs() {
    assert_eq!(
        run_prints(
            r#"<?php
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
"#
        ),
        &["0: Alice is 301: Bob is 25"]
    );
}

#[test]
fn generator_yields_map_values() {
    assert_eq!(
        run_prints(
            r#"<?php
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
"#
        ),
        &["Alice is 30Bob is 25"]
    );
}

// ── Generator valid/rewind ───────────────────────────────────────
#[test]
fn generator_valid_check() {
    assert_eq!(
        run_prints(
            r#"<?php
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
"#
        ),
        &["validdone"]
    );
}

// ── Generator as lazy infinite sequence ──────────────────────────
#[test]
fn generator_fibonacci_lazy() {
    assert_eq!(
        run_prints(
            r#"<?php
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
"#
        ),
        &["0,1,1,2,3,5,8,13"]
    );
}

// ── Generator take helper ────────────────────────────────────────
#[test]
fn generator_take_n() {
    assert_eq!(
        run_prints(
            r#"<?php
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
"#
        ),
        &["1,2,3,4,5"]
    );
}

// ── yield from return value ──────────────────────────────────────
#[test]
fn yield_from_captures_return() {
    assert_eq!(
        run_prints(
            r#"<?php
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
"#
        ),
        &["abdone"]
    );
}

// ── Generator with exception handling ────────────────────────────
#[test]
fn generator_throw() {
    assert_eq!(
        run_prints(
            r#"<?php
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
"#
        ),
        &["12caught: stop"]
    );
}

#[test]
fn generator_throw_before_first_yield() {
    assert_eq!(
        run_prints(
            r#"<?php
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
"#
        ),
        &["caughthandled"]
    );
}

#[test]
fn generator_variadic_throw_before_first_yield() {
    assert_eq!(
        run_prints(
            r#"<?php
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
"#
        ),
        &["b,cstop"]
    );
}

// ── Coroutine pattern ────────────────────────────────────────────
#[test]
fn coroutine_echo_back() {
    assert_eq!(
        run_prints(
            r#"<?php
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
"#
        ),
        &["hello,world"]
    );
}

// ── Generator with finally ───────────────────────────────────────
#[test]
fn generator_finally_cleanup() {
    assert_eq!(
        run_prints(
            r#"<?php
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
"#
        ),
        &["open12close"]
    );
}

// ── Generator as return type hint ────────────────────────────────
#[test]
fn generator_return_type_hint() {
    compile_ok(
        r#"<?php
function counter(int $start, int $end): Generator {
    for ($i = $start; $i <= $end; $i++) {
        yield $i;
    }
}
$g = counter(1, 3);
foreach ($g as $v) {
    echo $v;
}
"#,
    );
}

// ── yield from plain array delegation ───────────────────────────
#[test]
fn yield_from_plain_array_delegation() {
    assert_eq!(
        run_prints(
            r#"<?php
function fromArray(array $items) {
    yield from $items;
}
$result = [];
foreach (fromArray([10, 20, 30]) as $v) {
    $result[] = $v;
}
echo implode(",", $result);
"#
        ),
        &["10,20,30"]
    );
}

// ── Prime number sieve generator ─────────────────────────────────
#[test]
fn generator_prime_sieve() {
    assert_eq!(
        run_prints(
            r#"<?php
function primes(int $limit) {
    $sieve = array_fill(2, $limit - 1, true);
    for ($i = 2; $i <= $limit; $i++) {
        if (!isset($sieve[$i]) || !$sieve[$i]) continue;
        yield $i;
        for ($j = $i * 2; $j <= $limit; $j += $i) {
            $sieve[$j] = false;
        }
    }
}
$result = [];
foreach (primes(30) as $p) {
    $result[] = $p;
}
echo implode(",", $result);
"#
        ),
        &["2,3,5,7,11,13,17,19,23,29"]
    );
}

// ── Lazy range generator ─────────────────────────────────────────
#[test]
fn lazy_range_generator_memory_efficient() {
    assert_eq!(
        run_prints(
            r#"<?php
function lazyRange(float $start, float $end, float $step = 1.0) {
    for ($i = $start; $i <= $end; $i += $step) {
        yield $i;
    }
}
$result = [];
foreach (lazyRange(0, 1, 0.25) as $v) {
    $result[] = $v;
}
echo implode(",", $result);
"#
        ),
        &["0,0.25,0.5,0.75,1"]
    );
}

// ── Generator converted to array ─────────────────────────────────
#[test]
fn generator_to_array_via_iterator_to_array() {
    assert_eq!(
        run_prints(
            r#"<?php
function squares(int $n) {
    for ($i = 1; $i <= $n; $i++) {
        yield $i => $i * $i;
    }
}
$arr = iterator_to_array(squares(5));
echo implode(",", $arr);
"#
        ),
        &["1,4,9,16,25"]
    );
}

// ── Generator can't rewind ────────────────────────────────────────
#[test]
fn generator_rewind_no_op_after_start() {
    assert_eq!(
        run_prints(
            r#"<?php
function once() {
    yield "first";
    yield "second";
}
$g = once();
echo $g->current();
$g->next();
echo $g->current();
// rewind on a started generator is a no-op / throws; we just verify
// we can call it without crashing by wrapping in try/catch
try {
    $g->rewind();
} catch (Exception $e) {
    echo "rewind-error";
}
"#
        ),
        &["firstsecondrewind-error"]
    );
}

// ── Nested generators ────────────────────────────────────────────
#[test]
fn nested_generator_delegation() {
    assert_eq!(
        run_prints(
            r#"<?php
function level3() {
    yield "c";
    yield "d";
}
function level2() {
    yield "b";
    yield from level3();
    yield "e";
}
function level1() {
    yield "a";
    yield from level2();
    yield "f";
}
$result = [];
foreach (level1() as $v) {
    $result[] = $v;
}
echo implode("", $result);
"#
        ),
        &["abcdef"]
    );
}

// ── Generator in recursive function ──────────────────────────────
#[test]
fn generator_in_recursive_function() {
    assert_eq!(
        run_prints(
            r#"<?php
function permutations(array $items): Generator {
    if (count($items) <= 1) {
        yield $items;
        return;
    }
    foreach ($items as $k => $v) {
        $rest = $items;
        array_splice($rest, $k, 1);
        foreach (permutations($rest) as $perm) {
            yield array_merge([$v], $perm);
        }
    }
}
$count = 0;
foreach (permutations([1, 2, 3]) as $perm) {
    $count++;
}
echo $count; // 3! = 6
"#
        ),
        &["6"]
    );
}

// ── Generator yielding objects ────────────────────────────────────
#[test]
fn generator_yields_objects() {
    assert_eq!(
        run_prints(
            r#"<?php
class Point {
    public function __construct(public int $x, public int $y) {}
}
function points(array $coords) {
    foreach ($coords as [$x, $y]) {
        yield new Point($x, $y);
    }
}
$result = [];
foreach (points([[1,2],[3,4],[5,6]]) as $p) {
    $result[] = "({$p->x},{$p->y})";
}
echo implode(" ", $result);
"#
        ),
        &["(1,2) (3,4) (5,6)"]
    );
}

// ── Generator with complex state machine ─────────────────────────
#[test]
fn generator_state_machine_lexer() {
    assert_eq!(
        run_prints(
            r#"<?php
function tokenize(string $input) {
    $len = strlen($input);
    $i = 0;
    while ($i < $len) {
        if (ctype_space($input[$i])) { $i++; continue; }
        if (ctype_digit($input[$i])) {
            $num = "";
            while ($i < $len && ctype_digit($input[$i])) {
                $num .= $input[$i++];
            }
            yield ["NUM", $num];
        } elseif (str_contains("+-*/", $input[$i])) {
            yield ["OP", $input[$i]];
            $i++;
        } else {
            yield ["UNK", $input[$i]];
            $i++;
        }
    }
}
$tokens = [];
foreach (tokenize("12 + 34 * 5") as [$type, $val]) {
    $tokens[] = "$type:$val";
}
echo implode("|", $tokens);
"#
        ),
        &["NUM:12|OP:+|NUM:34|OP:*|NUM:5"]
    );
}

// ── send() return value is yield expression value ────────────────
#[test]
fn send_return_is_yield_expression_value() {
    assert_eq!(
        run_prints(
            r#"<?php
function doubler() {
    while (true) {
        $input = yield;
        if ($input === null) return;
        yield $input * 2;
    }
}
$g = doubler();
$g->current(); // prime
echo $g->send(5);  // yields 10
$g->next();        // advance to next yield
echo $g->send(7);  // yields 14
"#
        ),
        &["1014"]
    );
}

// ── Generator tracking external state via use ────────────────────
#[test]
fn generator_tracks_external_state() {
    assert_eq!(
        run_prints(
            r#"<?php
function makeTracker(array &$log) {
    return (function() use (&$log) {
        $log[] = "started";
        yield 1;
        $log[] = "middle";
        yield 2;
        $log[] = "ended";
    })();
}
$log = [];
$gen = makeTracker($log);
foreach ($gen as $v) {
    // consume
}
echo implode(",", $log);
"#
        ),
        &["started,middle,ended"]
    );
}

// ── Generator producing no values ────────────────────────────────
#[test]
fn generator_empty_produces_no_values() {
    assert_eq!(
        run_prints(
            r#"<?php
function empty_gen() {
    return;
    yield; // unreachable but makes it a generator
}
$g = empty_gen();
echo $g->valid() ? "has values" : "empty";
"#
        ),
        &["empty"]
    );
}

// ── Generator with default parameter ─────────────────────────────
#[test]
fn generator_with_default_parameter() {
    assert_eq!(
        run_prints(
            r#"<?php
function counter(int $start = 0, int $step = 1) {
    $n = $start;
    while (true) {
        yield $n;
        $n += $step;
    }
}
$g = counter(10, 5);
$result = [];
for ($i = 0; $i < 4; $i++) {
    $result[] = $g->current();
    $g->next();
}
echo implode(",", $result);
"#
        ),
        &["10,15,20,25"]
    );
}

// ── yield null explicitly ────────────────────────────────────────
#[test]
fn generator_yield_null_explicitly() {
    assert_eq!(
        run_prints(
            r#"<?php
function nulls(int $count) {
    for ($i = 0; $i < $count; $i++) {
        yield null;
    }
}
$c = 0;
foreach (nulls(3) as $v) {
    echo $v === null ? "null" : "not-null";
    $c++;
}
echo $c;
"#
        ),
        &["nullnullnull3"]
    );
}

// ── Zip two generators together ──────────────────────────────────
#[test]
fn zip_two_generators() {
    assert_eq!(
        run_prints(
            r#"<?php
function zipGens(Generator $a, Generator $b) {
    while ($a->valid() && $b->valid()) {
        yield [$a->current(), $b->current()];
        $a->next();
        $b->next();
    }
}
function letters() {
    foreach (["a", "b", "c"] as $l) yield $l;
}
function numbers() {
    foreach ([1, 2, 3] as $n) yield $n;
}
$result = [];
foreach (zipGens(letters(), numbers()) as [$l, $n]) {
    $result[] = "$l$n";
}
echo implode(",", $result);
"#
        ),
        &["a1,b2,c3"]
    );
}

// ── Generator while valid() + send() ────────────────────────────
#[test]
fn generator_while_valid_with_send() {
    assert_eq!(
        run_prints(
            r#"<?php
function multiplier() {
    $factor = yield "ready";
    while (true) {
        $value = yield;
        if ($value === null) return;
        yield $value * $factor;
    }
}
$g = multiplier();
$g->current(); // "ready"
$g->send(3);   // sets factor, yields null
$g->next();    // advance to inner yield
echo $g->send(5);   // yields 15
$g->next();
echo $g->send(7);   // yields 21
"#
        ),
        &["15", "21"]
    );
}

// ── Generator lazy evaluation vs array eager ─────────────────────
#[test]
fn generator_lazy_vs_eager_evaluation() {
    assert_eq!(
        run_prints(
            r#"<?php
$calls_gen = 0;
$calls_arr = 0;
function lazyDoubles(array $items, int &$counter) {
    foreach ($items as $v) {
        $counter++;
        yield $v * 2;
    }
}
function eagerDoubles(array $items, int &$counter): array {
    $counter += count($items);
    return array_map(fn($v) => $v * 2, $items);
}
$data = [1, 2, 3, 4, 5];
$gen = lazyDoubles($data, $calls_gen);
// only consume first 2
$taken = [];
for ($i = 0; $i < 2; $i++) {
    $taken[] = $gen->current();
    $gen->next();
}
echo $calls_gen;     // only 2 calls made
$arr = eagerDoubles($data, $calls_arr);
echo $calls_arr;     // all 5 evaluated upfront
echo implode(",", $taken);
"#
        ),
        &["352,4"]
    );
}

// ── yield in match expression ─────────────────────────────────────
#[test]
fn yield_in_match_expression() {
    assert_eq!(
        run_prints(
            r#"<?php
function classifyNumbers(array $nums) {
    foreach ($nums as $n) {
        yield match(true) {
            $n < 0  => "neg",
            $n === 0 => "zero",
            $n > 0  => "pos",
        };
    }
}
$result = [];
foreach (classifyNumbers([-2, 0, 3, -1, 5]) as $label) {
    $result[] = $label;
}
echo implode(",", $result);
"#
        ),
        &["neg,zero,pos,neg,pos"]
    );
}

// ── yield after complex expression ───────────────────────────────
#[test]
fn yield_after_complex_expression() {
    assert_eq!(
        run_prints(
            r#"<?php
function transformedRange(int $n) {
    for ($i = 1; $i <= $n; $i++) {
        yield $i % 2 === 0
            ? $i * $i
            : $i * 2 + 1;
    }
}
$result = [];
foreach (transformedRange(6) as $v) {
    $result[] = $v;
}
echo implode(",", $result);
// i=1 odd: 1*2+1=3, i=2 even: 4, i=3 odd: 7, i=4 even: 16, i=5 odd: 11, i=6 even: 36
"#
        ),
        &["3,4,7,16,11,36"]
    );
}

// ── Generator infinite loop broken by consumer ───────────────────
#[test]
fn generator_infinite_loop_broken_by_consumer() {
    assert_eq!(
        run_prints(
            r#"<?php
function incrementsForever(int $start = 0) {
    $n = $start;
    while (true) {
        yield $n++;
    }
}
$g = incrementsForever(1);
$result = [];
foreach ($g as $v) {
    $result[] = $v;
    if ($v >= 5) break;
}
echo implode(",", $result);
"#
        ),
        &["1,2,3,4,5"]
    );
}

// ── Generator composition pipeline via yield from ─────────────────
#[test]
fn generator_composition_chained_pipeline() {
    let result = std::thread::Builder::new()
        .stack_size(8 * 1024 * 1024)
        .spawn(|| {
            run_prints(
                r#"<?php
function source(array $items) {
    yield from $items;
}
function squared($gen) {
    foreach ($gen as $v) yield $v * $v;
}
function onlyOdd($gen) {
    foreach ($gen as $v) {
        if ($v % 2 !== 0) yield $v;
    }
}
function asStrings($gen) {
    foreach ($gen as $v) yield (string)$v;
}
$pipeline = asStrings(onlyOdd(squared(source([1, 2, 3, 4, 5]))));
echo implode(",", iterator_to_array($pipeline, false));
"#,
            )
        })
        .unwrap()
        .join()
        .unwrap();
    assert_eq!(result, &["1,9,25"]);
}

// ── Multiple generators running interleaved via manual dispatch ──
#[test]
fn multiple_generators_interleaved_round_robin() {
    assert_eq!(
        run_prints(
            r#"<?php
function taskA() { yield "A1"; yield "A2"; yield "A3"; }
function taskB() { yield "B1"; yield "B2"; }
$gens = [taskA(), taskB()];
$output = [];
$alive = true;
while ($alive) {
    $alive = false;
    foreach ($gens as $g) {
        if ($g->valid()) {
            $output[] = $g->current();
            $g->next();
            $alive = true;
        }
    }
}
echo implode(",", $output);
"#
        ),
        &["A1,B1,A2,B2,A3"]
    );
}

// ── Generator preserving reference between yields ────────────────
#[test]
fn generator_preserves_reference_between_yields() {
    assert_eq!(
        run_prints(
            r#"<?php
function buildString() {
    $parts = [];
    while (true) {
        $part = yield implode("", $parts);
        if ($part === null) break;
        $parts[] = $part;
    }
}
$g = buildString();
$g->current(); // start with empty
$g->send("foo");
$g->send("bar");
echo $g->send("baz");
"#
        ),
        &["foobarbaz"]
    );
}

// ── Generator return type in interface method ────────────────────
#[test]
fn generator_return_type_in_interface() {
    compile_ok(
        r#"<?php
interface Iterable2 {
    public function items(): Generator;
}
class NumberList implements Iterable2 {
    private array $data;
    public function __construct(array $data) { $this->data = $data; }
    public function items(): Generator {
        foreach ($this->data as $v) {
            yield $v;
        }
    }
}
$list = new NumberList([10, 20, 30]);
foreach ($list->items() as $v) {
    echo $v;
}
"#,
    );
}

// ── iterator_to_array with preserve_keys false ───────────────────
#[test]
fn iterator_to_array_no_preserve_keys() {
    assert_eq!(
        run_prints(
            r#"<?php
function words() {
    yield 5 => "apple";
    yield 3 => "banana";
    yield 7 => "cherry";
}
$arr = iterator_to_array(words(), false);
echo implode(",", $arr);
"#
        ),
        &["apple,banana,cherry"]
    );
}
