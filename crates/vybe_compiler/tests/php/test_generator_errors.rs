//! Generator exception/resume paths not covered by send_throw/advanced suites.

crate::php_cases! {
    foreach_stops_when_generator_throws_on_second_yield => {
        r#"<?php
function gen(): Generator {
    yield 1;
    throw new RuntimeException('stop');
    yield 2;
}
$log = [];
try {
    foreach (gen() as $v) { $log[] = $v; }
} catch (RuntimeException $e) {
    $log[] = 'caught';
}
echo implode(',', $log);
"#,
        ["1,caught"]
    };

    yield_from_propagates_inner_throw_to_outer_catch => {
        r#"<?php
function inner(): Generator { throw new LogicException('inner'); yield 1; }
function outer(): Generator { yield from inner(); }
try { iterator_to_array(outer()); echo 'ok'; }
catch (LogicException $e) { echo $e->getMessage(); }
"#,
        ["inner"]
    };

    yield_from_outer_try_catch_around_expression => {
        r#"<?php
function ok(): Generator { yield 'a'; }
function bad(): Generator { throw new Exception('x'); yield 'b'; }
function runner(): Generator {
    try { yield from bad(); }
    catch (Exception $e) { yield 'recovered'; }
}
echo implode('', iterator_to_array(runner()));
"#,
        ["recovered"]
    };

    generator_get_return_after_normal_completion => {
        r#"<?php
function finish(): Generator {
    yield 1;
    return 'done';
}
$g = finish();
$g->next();
$g->next();
echo $g->getReturn();
"#,
        ["done"]
    };

    generator_cannot_get_return_before_done => {
        r#"<?php
function g(): Generator { yield 1; return 2; }
$gen = g();
$gen->next();
try { echo $gen->getReturn(); }
catch (Exception $e) { echo 'early'; }
"#,
        ["early"]
    };

    generator_valid_false_after_completion => {
        r#"<?php
function one(): Generator { yield 'only'; }
$g = one();
$g->next();
$g->next();
echo $g->valid() ? 'valid' : 'done';
"#,
        ["done"]
    };

    generator_current_after_completion_is_null => {
        r#"<?php
function one(): Generator { yield 'x'; }
$g = one();
$g->next();
$g->next();
echo $g->current() === null ? 'null' : 'set';
"#,
        ["null"]
    };

    generator_throw_then_resume_is_invalid => {
        r#"<?php
function g(): Generator { yield 1; }
$gen = g();
$gen->next();
try {
    $gen->throw(new RuntimeException('boom'));
    $gen->next();
    echo 'resumed';
} catch (RuntimeException $e) {
    echo 'thrown';
}
"#,
        ["thrown"]
    };

    generator_send_after_throw_attempt_fails => {
        r#"<?php
function acc(): Generator {
    $total = 0;
    while (true) {
        $n = yield $total;
        if ($n === null) break;
        $total += $n;
    }
}
$g = acc();
$g->current();
try {
    $g->throw(new InvalidArgumentException('abort'));
    $g->send(5);
    echo 'sent';
} catch (InvalidArgumentException $e) {
    echo 'abort';
}
"#,
        ["abort"]
    };

    nested_generators_bubble_throw_from_delegate => {
        r#"<?php
function leaf(): Generator { throw new DomainException('leaf'); yield 0; }
function mid(): Generator { yield from leaf(); }
function top(): Generator { yield from mid(); }
try { foreach (top() as $_) {} echo 'ok'; }
catch (DomainException $e) { echo 'leaf'; }
"#,
        ["leaf"]
    };

    generator_finally_runs_after_throw_in_try => {
        r#"<?php
function g(): Generator {
    try { throw new RuntimeException('t'); yield 1; }
    finally { yield 'f'; }
}
$log = [];
try { foreach (g() as $v) { $log[] = $v; } }
catch (RuntimeException $e) { $log[] = 'c'; }
echo implode('', $log);
"#,
        ["c"]
    };

    generator_return_value_not_yielded_to_foreach => {
        r#"<?php
function withReturn(): Generator {
    yield 'y';
    return 'ret';
}
$seen = [];
foreach (withReturn() as $v) { $seen[] = $v; }
echo implode(',', $seen);
"#,
        ["y"]
    };

    generator_keys_preserved_in_foreach => {
        r#"<?php
function keyed(): Generator {
    yield 'a' => 1;
    yield 'b' => 2;
}
$keys = [];
foreach (keyed() as $k => $v) { $keys[] = $k; }
echo implode('', $keys);
"#,
        ["ab"]
    };

    generator_yield_reference_mutates_outer => {
        r#"<?php
function byRef(): Generator {
    $x = 1;
    yield &$x;
    $x = 3;
}
$g = byRef();
$ref = &$g->current();
$g->next();
echo $ref;
"#,
        ["3"]
    };

    multiple_yield_from_sequence_concatenates => {
        r#"<?php
function a(): Generator { yield 'A'; }
function b(): Generator { yield 'B'; }
function both(): Generator { yield from a(); yield from b(); }
echo implode('', iterator_to_array(both()));
"#,
        ["AB"]
    };

    generator_manual_next_advances_value => {
        r#"<?php
function nums(): Generator { yield 2; yield 4; }
$it = nums();
$it->next();
echo $it->current();
$it->next();
echo $it->current();
"#,
        ["24"]
    };

    generator_rewind_not_supported => {
        r#"<?php
function g(): Generator { yield 1; }
$gen = g();
try { $gen->rewind(); echo 'rewound'; }
catch (Exception $e) { echo 'norewind'; }
"#,
        ["norewind"]
    };

    generator_with_empty_body_only_return => {
        r#"<?php
function emptyGen(): Generator { return; yield; }
$g = emptyGen();
echo $g->valid() ? 'valid' : 'invalid';
"#,
        ["invalid"]
    };

    generator_only_yield_null => {
        r#"<?php
function nulls(): Generator { yield null; }
$g = nulls();
$g->next();
echo $g->current() === null ? 'null' : 'val';
"#,
        ["null"]
    };

    generator_delegates_return_from_inner_via_yield_from => {
        r#"<?php
function inner(): Generator { yield 1; return 9; }
function outer(): Generator { return yield from inner(); }
$g = outer();
$g->next();
$g->next();
try { echo $g->getReturn(); } catch (Exception $e) { echo 'no'; }
"#,
        ["9"]
    };

    catch_inside_generator_resumes_after_exception => {
        r#"<?php
function resilient(): Generator {
    try { throw new RuntimeException('x'); }
    catch (RuntimeException $e) { yield 'caught'; }
    yield 'after';
}
echo implode(',', iterator_to_array(resilient()));
"#,
        ["caught,after"]
    };

    generator_throw_caught_by_outer_try_not_foreach => {
        r#"<?php
function boom(): Generator { yield 1; throw new Exception('e'); }
try {
    $g = boom();
    $g->next();
    $g->next();
} catch (Exception $e) {
    echo $e->getMessage();
}
"#,
        ["e"]
    };

    generator_foreach_break_before_throw => {
        r#"<?php
function long(): Generator {
    yield 1;
    yield 2;
    throw new RuntimeException('late');
}
$log = [];
foreach (long() as $v) {
    $log[] = $v;
    if ($v === 1) break;
}
echo implode(',', $log);
"#,
        ["1"]
    };

    generator_map_via_manual_next_loop => {
        r#"<?php
function double(): Generator {
    foreach ([1, 2, 3] as $n) { yield $n * 2; }
}
$out = [];
$g = double();
while ($g->valid()) {
    $out[] = $g->current();
    $g->next();
}
echo implode('-', $out);
"#,
        ["2-4-6"]
    };

    generator_closure_factory => {
        r#"<?php
$make = function(int $n): Generator {
    for ($i = 0; $i < $n; $i++) { yield $i; }
};
echo implode('', iterator_to_array($make(3)));
"#,
        ["012"]
    };

    generator_method_on_object => {
        r#"<?php
class Seq {
    public function run(): Generator { yield 'm'; }
}
echo implode('', iterator_to_array((new Seq())->run()));
"#,
        ["m"]
    };

    generator_static_method => {
        r#"<?php
class Seq {
    public static function run(): Generator { yield 's'; }
}
echo implode('', iterator_to_array(Seq::run()));
"#,
        ["s"]
    };

    generator_anonymous_function => {
        r#"<?php
$gen = function(): Generator { yield 'anon'; };
echo implode('', iterator_to_array($gen()));
"#,
        ["anon"]
    };

    generator_passed_to_iterator_apply => {
        r#"<?php
function g(): Generator { yield 1; yield 2; }
$sum = 0;
iterator_apply(g(), function($it) use (&$sum) {
    foreach ($it as $v) { $sum += $v; }
});
echo $sum;
"#,
        ["3"]
    };

    generator_with_try_finally_yield_order => {
        r#"<?php
function order(): Generator {
    try { yield 't'; }
    finally { yield 'f'; }
}
echo implode('', iterator_to_array(order()));
"#,
        ["tf"]
    };

    generator_delegate_empty_inner => {
        r#"<?php
function emptyInner(): Generator { return; yield; }
function outer(): Generator { yield 'o'; yield from emptyInner(); yield 'p'; }
echo implode('', iterator_to_array(outer()));
"#,
        ["op"]
    };

    generator_exception_in_condition_before_yield => {
        r#"<?php
function check(bool $ok): Generator {
    if (!$ok) { throw new InvalidArgumentException('bad'); }
    yield 'ok';
}
try { echo implode('', iterator_to_array(check(false))); }
catch (InvalidArgumentException $e) { echo 'bad'; }
"#,
        ["bad"]
    };

    generator_yield_object_by_value => {
        r#"<?php
function obj(): Generator { yield new ArrayObject([1]); }
$g = obj();
$g->next();
echo $g->current() instanceof ArrayObject ? 'obj' : 'no';
"#,
        ["obj"]
    };

    generator_multiple_throw_attempts_only_first_matters => {
        r#"<?php
function g(): Generator { yield 1; }
$gen = g();
$gen->next();
try {
    $gen->throw(new RuntimeException('one'));
} catch (RuntimeException $e) {
    echo $e->getMessage();
}
"#,
        ["one"]
    };

    generator_from_array_walk => {
        r#"<?php
$acc = [];
$gen = (function() use (&$acc) {
    foreach ([1, 2] as $v) { $acc[] = $v; yield $v; }
})();
iterator_to_array($gen);
echo implode('+', $acc);
"#,
        ["1+2"]
    };

    generator_to_array_preserves_keys_false => {
        r#"<?php
function g(): Generator { yield 'k' => 9; }
$arr = iterator_to_array(g(), false);
echo implode('', array_keys($arr));
"#,
        ["k"]
    };

    generator_to_array_preserves_keys_true => {
        r#"<?php
function g(): Generator { yield 5 => 'v'; }
$arr = iterator_to_array(g(), true);
echo $arr[5];
"#,
        ["v"]
    };

    generator_fibonacci_style_state => {
        r#"<?php
function fib(int $n): Generator {
    $a = 0; $b = 1;
    for ($i = 0; $i < $n; $i++) {
        yield $a;
        [$a, $b] = [$b, $a + $b];
    }
}
echo implode(',', iterator_to_array(fib(4)));
"#,
        ["0,1,1,2"]
    };

    generator_cleanup_via_unset => {
        r#"<?php
function g(): Generator { yield 1; }
$gen = g();
$gen->next();
unset($gen);
echo 'unset';
"#,
        ["unset"]
    };

    generator_class_implements_traversable => {
        r#"<?php
function g(): Generator { yield 1; }
$gen = g();
echo $gen instanceof Traversable ? 'traversable' : 'no';
"#,
        ["traversable"]
    };

    generator_class_implements_iterator => {
        r#"<?php
function g(): Generator { yield 1; }
$gen = g();
echo $gen instanceof Iterator ? 'iterator' : 'no';
"#,
        ["iterator"]
    };

    generator_yield_stringable_cast_in_concat => {
        r#"<?php
class Label { public function __construct(private string $t) {} public function __toString(): string { return $this->t; } }
function g(): Generator { yield new Label('z'); }
$g = g();
$g->next();
echo (string)$g->current();
"#,
        ["z"]
    };

    generator_nested_yield_from_with_return_in_middle => {
        r#"<?php
function mid(): Generator { yield 'm'; return 'R'; }
function top(): Generator { yield 't'; $r = yield from mid(); yield $r; }
echo implode(',', iterator_to_array(top()));
"#,
        ["t,m,R"]
    };
}
