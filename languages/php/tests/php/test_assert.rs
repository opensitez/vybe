//! Runtime assert() and assert_options — zero overlap with compile-only error tests.

crate::php_cases! {
    assert_true_expression_allows_following_echo => {
        r#"<?php
assert(1 + 1 === 2);
echo 'ok';
"#,
        ["ok"]
    };

    assert_false_with_exception_mode_yields_assertion_error => {
        r#"<?php
assert_options(ASSERT_EXCEPTION, 1);
try { assert(false); echo 'no'; }
catch (AssertionError $e) { echo 'assertion'; }
"#,
        ["assertion"]
    };

    assert_description_appended_to_message => {
        r#"<?php
assert_options(ASSERT_EXCEPTION, 1);
try { assert(false, 'bad state'); }
catch (AssertionError $e) { echo $e->getMessage(); }
"#,
        ["bad state"]
    };

    assert_inactive_skips_failing_expression => {
        r#"<?php
$prev = assert_options(ASSERT_ACTIVE, 0);
assert(false);
assert_options(ASSERT_ACTIVE, $prev);
echo 'skipped';
"#,
        ["skipped"]
    };

    assert_callback_receives_assertion_details => {
        r#"<?php
$seen = '';
$old = assert_options(ASSERT_CALLBACK, function($file, $line, $assertion, $desc = null) use (&$seen) {
    $seen = $line > 0 ? 'cb' : 'no';
    return true;
});
assert(false, 'via callback');
assert_options(ASSERT_CALLBACK, $old);
echo $seen;
"#,
        ["cb"]
    };

    assert_callback_return_true_suppresses_default => {
        r#"<?php
$old = assert_options(ASSERT_CALLBACK, fn() => true);
assert(false);
assert_options(ASSERT_CALLBACK, $old);
echo 'handled';
"#,
        ["handled"]
    };

    assert_options_restore_exception_flag => {
        r#"<?php
$was = assert_options(ASSERT_EXCEPTION, 1);
$now = assert_options(ASSERT_EXCEPTION);
assert_options(ASSERT_EXCEPTION, $was);
echo $now === 1 ? 'on' : 'off';
"#,
        ["on"]
    };

    assert_inside_nested_function_bubbles => {
        r#"<?php
assert_options(ASSERT_EXCEPTION, 1);
function check(bool $ok): void { assert($ok); }
try { check(false); echo 'no'; }
catch (AssertionError $e) { echo 'nested'; }
"#,
        ["nested"]
    };

    assert_inside_static_method => {
        r#"<?php
assert_options(ASSERT_EXCEPTION, 1);
class Guard {
    public static function must(bool $ok): void { assert($ok); }
}
try { Guard::must(false); echo 'no'; }
catch (AssertionError $e) { echo 'static'; }
"#,
        ["static"]
    };

    assert_inside_closure => {
        r#"<?php
assert_options(ASSERT_EXCEPTION, 1);
$run = function(bool $ok) { assert($ok); };
try { $run(false); echo 'no'; }
catch (AssertionError $e) { echo 'closure'; }
"#,
        ["closure"]
    };

    assert_after_successful_try_catch_continues => {
        r#"<?php
assert_options(ASSERT_EXCEPTION, 1);
try { throw new Exception('x'); }
catch (Exception $e) { /* handled */ }
assert(true);
echo 'after';
"#,
        ["after"]
    };

    assert_with_array_count_expression_true => {
        r#"<?php
assert(count([1, 2, 3]) === 3);
echo 'arr';
"#,
        ["arr"]
    };

    assert_with_array_count_expression_false => {
        r#"<?php
assert_options(ASSERT_EXCEPTION, 1);
try { assert(count([]) > 0); }
catch (AssertionError $e) { echo 'empty'; }
"#,
        ["empty"]
    };

    assert_combined_with_boolean_and_short_circuit => {
        r#"<?php
$flag = false;
assert($flag || true);
echo 'and';
"#,
        ["and"]
    };

    assert_combined_with_boolean_or_triggers_on_both_false => {
        r#"<?php
assert_options(ASSERT_EXCEPTION, 1);
try { assert(false || false); }
catch (AssertionError $e) { echo 'or'; }
"#,
        ["or"]
    };

    assert_on_string_equality => {
        r#"<?php
assert('php' === 'php');
echo 'str';
"#,
        ["str"]
    };

    assert_on_string_inequality_fails => {
        r#"<?php
assert_options(ASSERT_EXCEPTION, 1);
try { assert('a' === 'b', 'mismatch'); }
catch (AssertionError $e) { echo 'neq'; }
"#,
        ["neq"]
    };

    assert_on_is_int_type_check => {
        r#"<?php
assert(is_int(7));
echo 'int';
"#,
        ["int"]
    };

    assert_on_is_array_type_check_false => {
        r#"<?php
assert_options(ASSERT_EXCEPTION, 1);
try { assert(is_array(3)); }
catch (AssertionError $e) { echo 'not array'; }
"#,
        ["not array"]
    };

    assert_on_is_null_check => {
        r#"<?php
assert(is_null(null));
echo 'null';
"#,
        ["null"]
    };

    assert_on_is_object_check => {
        r#"<?php
assert(is_object(new stdClass()));
echo 'obj';
"#,
        ["obj"]
    };

    assert_on_is_callable_check => {
        r#"<?php
assert(is_callable('strlen'));
echo 'callable';
"#,
        ["callable"]
    };

    assert_on_in_array_membership => {
        r#"<?php
assert(in_array(2, [1, 2, 3], true));
echo 'in';
"#,
        ["in"]
    };

    assert_on_in_array_strict_miss => {
        r#"<?php
assert_options(ASSERT_EXCEPTION, 1);
try { assert(in_array('1', [1, 2, 3], true)); }
catch (AssertionError $e) { echo 'miss'; }
"#,
        ["miss"]
    };

    assert_on_array_key_exists => {
        r#"<?php
assert(array_key_exists('a', ['a' => 1]));
echo 'key';
"#,
        ["key"]
    };

    assert_on_array_key_missing => {
        r#"<?php
assert_options(ASSERT_EXCEPTION, 1);
try { assert(array_key_exists('z', ['a' => 1])); }
catch (AssertionError $e) { echo 'nokey'; }
"#,
        ["nokey"]
    };

    assert_on_match_expression_result => {
        r#"<?php
assert(match (2) { 1 => false, 2 => true, default => false });
echo 'match';
"#,
        ["match"]
    };

    assert_on_null_coalesce_defined => {
        r#"<?php
$v = ['k' => 1];
assert(($v['k'] ?? null) === 1);
echo 'coalesce';
"#,
        ["coalesce"]
    };

    assert_on_spaceship_zero => {
        r#"<?php
assert((3 <=> 3) === 0);
echo 'spaceship';
"#,
        ["spaceship"]
    };

    assert_on_bitwise_and_mask => {
        r#"<?php
assert((5 & 1) === 1);
echo 'bit';
"#,
        ["bit"]
    };

    assert_on_instanceof_check => {
        r#"<?php
assert((new Exception()) instanceof Exception);
echo 'instanceof';
"#,
        ["instanceof"]
    };

    assert_on_enum_case_equality => {
        r#"<?php
enum Color { case Red; case Blue; }
assert(Color::Red === Color::Red);
echo 'enum';
"#,
        ["enum"]
    };

    assert_caught_as_throwable_not_exception => {
        r#"<?php
assert_options(ASSERT_EXCEPTION, 1);
try { assert(false); }
catch (Throwable $t) { echo $t instanceof AssertionError ? 'throwable' : 'other'; }
"#,
        ["throwable"]
    };

    assert_does_not_extend_exception_hierarchy => {
        r#"<?php
assert_options(ASSERT_EXCEPTION, 1);
try { assert(false); }
catch (Exception $e) { echo 'exception'; }
catch (AssertionError $e) { echo 'assertion error'; }
"#,
        ["assertion error"]
    };

    assert_in_foreach_guard => {
        r#"<?php
assert_options(ASSERT_EXCEPTION, 1);
$log = [];
foreach ([1, 0, 2] as $n) {
    try {
        assert($n !== 0);
        $log[] = (string)$n;
    } catch (AssertionError $e) {
        $log[] = 'fail';
    }
}
echo implode(',', $log);
"#,
        ["1,fail,2"]
    };

    assert_in_while_loop_break_on_failure => {
        r#"<?php
assert_options(ASSERT_EXCEPTION, 1);
$i = 0;
$out = '';
while ($i < 3) {
    try {
        assert($i !== 1);
        $out .= $i;
    } catch (AssertionError $e) {
        $out .= 'X';
        break;
    }
    $i++;
}
echo $out;
"#,
        ["0X"]
    };

    assert_with_stringable_object_cast => {
        r#"<?php
class Label { public function __construct(private string $t) {} public function __toString(): string { return $this->t; } }
assert((string)new Label('x') === 'x');
echo 'stringable';
"#,
        ["stringable"]
    };

    assert_on_json_decode_valid => {
        r#"<?php
$data = json_decode('{"ok":true}', true);
assert($data['ok'] === true);
echo 'json';
"#,
        ["json"]
    };

    assert_on_preg_match_success => {
        r#"<?php
assert(preg_match('/^abc$/', 'abc') === 1);
echo 'preg';
"#,
        ["preg"]
    };

    assert_on_preg_match_failure => {
        r#"<?php
assert_options(ASSERT_EXCEPTION, 1);
try { assert(preg_match('/^abc$/', 'ab') === 1); }
catch (AssertionError $e) { echo 'preg fail'; }
"#,
        ["preg fail"]
    };

    assert_on_class_constant_equality => {
        r#"<?php
class Codes { public const OK = 200; }
assert(Codes::OK === 200);
echo 'const';
"#,
        ["const"]
    };

    assert_on_readonly_property_value => {
        r#"<?php
readonly class Id { public function __construct(public int $v) {} }
$id = new Id(9);
assert($id->v === 9);
echo 'readonly';
"#,
        ["readonly"]
    };

    assert_on_generator_valid_state => {
        r#"<?php
function one(): Generator { yield 1; }
$g = one();
assert($g->valid());
echo 'gen';
"#,
        ["gen"]
    };

    assert_on_fiber_not_started => {
        r#"<?php
$fiber = new Fiber(function(): void { Fiber::suspend('x'); });
assert(!$fiber->isStarted());
echo 'fiber';
"#,
        ["fiber"]
    };

    assert_options_bail_flag_readable => {
        r#"<?php
$val = assert_options(ASSERT_BAIL);
echo is_int($val) ? 'bail-int' : 'other';
"#,
        ["bail-int"]
    };

    assert_options_warning_flag_readable => {
        r#"<?php
$val = assert_options(ASSERT_WARNING);
echo is_int($val) ? 'warn-int' : 'other';
"#,
        ["warn-int"]
    };

    assert_options_quiet_eval_flag_readable => {
        r#"<?php
$val = assert_options(ASSERT_QUIET_EVAL);
echo is_int($val) ? 'quiet-int' : 'other';
"#,
        ["quiet-int"]
    };

    assert_in_trait_method => {
        r#"<?php
trait Checker { public function ok(): void { assert(true); } }
class App { use Checker; }
(new App())->ok();
echo 'trait';
"#,
        ["trait"]
    };

    assert_in_interface_default_method => {
        r#"<?php
interface Verifier {
    public function check(): void { assert(1 < 2); }
}
class Impl implements Verifier {}
(new Impl())->check();
echo 'iface';
"#,
        ["iface"]
    };

    assert_on_list_destructure_count => {
        r#"<?php
[$a, $b] = [1, 2];
assert($a + $b === 3);
echo 'list';
"#,
        ["list"]
    };

    assert_on_named_argument_call => {
        r#"<?php
function add(int $a, int $b): int { return $a + $b; }
assert(add(a: 2, b: 5) === 7);
echo 'named';
"#,
        ["named"]
    };

    assert_on_first_class_callable_invoke => {
        r#"<?php
function double(int $n): int { return $n * 2; }
$fn = double(...);
assert($fn(3) === 6);
echo 'fcc';
"#,
        ["fcc"]
    };

    assert_on_spl_queue_count => {
        r#"<?php
$q = new SplQueue();
$q->enqueue(1);
$q->enqueue(2);
assert($q->count() === 2);
echo 'queue';
"#,
        ["queue"]
    };
}
