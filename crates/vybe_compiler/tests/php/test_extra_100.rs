use super::helpers::run_prints;

// ── String functions ──────────────────────────────────────────

#[test]
fn sprintf_percent_literal() {
    assert_eq!(run_prints(r#"<?php echo sprintf('100%%'); "#), vec!["100%"]);
}
#[test]
fn str_split_empty_returns_false() {
    assert_eq!(
        run_prints(r#"<?php var_export(str_split('', 1) === []); "#),
        vec!["true"]
    );
}
#[test]
fn substr_out_of_range_empty() {
    assert_eq!(
        run_prints(r#"<?php var_export(substr('hello', 10)); "#),
        vec!["false"]
    );
}
#[test]
fn str_pad_both_odd_extra_right() {
    assert_eq!(
        run_prints(r#"<?php echo str_pad('x', 6, '-', STR_PAD_BOTH); "#),
        vec!["--x---"]
    );
}
#[test]
fn strstr_before_needle() {
    assert_eq!(
        run_prints(r#"<?php echo strstr('user@example.com', '@', true); "#),
        vec!["user"]
    );
}
#[test]
fn strstr_after_needle() {
    assert_eq!(
        run_prints(r#"<?php echo strstr('user@example.com', '@'); "#),
        vec!["@example.com"]
    );
}
#[test]
fn str_contains_empty_needle_always_true() {
    assert_eq!(
        run_prints(r#"<?php echo str_contains('', '') ? 'yes' : 'no'; "#),
        vec!["yes"]
    );
}
#[test]
fn number_format_rounding() {
    assert_eq!(
        run_prints(r#"<?php echo number_format(1.005, 2); "#),
        vec!["1.01"]
    );
}
#[test]
fn implode_single_element() {
    assert_eq!(
        run_prints(r#"<?php echo implode(',', ['alone']); "#),
        vec!["alone"]
    );
}
#[test]
fn explode_no_match_returns_single() {
    assert_eq!(
        run_prints(r#"<?php echo count(explode(',', 'hello')); "#),
        vec!["1"]
    );
}

// ── Array functions ───────────────────────────────────────────

#[test]
fn array_pop_empty_returns_null() {
    assert_eq!(
        run_prints(r#"<?php $a = []; var_export(array_pop($a)); "#),
        vec!["NULL"]
    );
}
#[test]
fn array_shift_empty_returns_null() {
    assert_eq!(
        run_prints(r#"<?php $a = []; var_export(array_shift($a)); "#),
        vec!["NULL"]
    );
}
#[test]
fn array_merge_with_empty() {
    assert_eq!(
        run_prints(r#"<?php echo implode(',', array_merge([], [1,2], [])); "#),
        vec!["1,2"]
    );
}
#[test]
fn array_keys_with_search_value() {
    assert_eq!(
        run_prints(r#"<?php echo implode(',', array_keys([1,2,1,3,1], 1)); "#),
        vec!["0,2,4"]
    );
}
#[test]
fn array_values_reindexes() {
    assert_eq!(
        run_prints(
            r#"<?php $a = [5=>'x',2=>'y']; echo implode(',', array_keys(array_values($a))); "#
        ),
        vec!["0,1"]
    );
}
#[test]
fn in_array_loose() {
    assert_eq!(
        run_prints(r#"<?php echo in_array('1', [1, 2, 3]) ? 'yes' : 'no'; "#),
        vec!["yes"]
    );
}
#[test]
fn array_count_values_string_keys() {
    assert_eq!(
        run_prints(r#"<?php $c = array_count_values(['a','b','a']); echo $c['a']; "#),
        vec!["2"]
    );
}
#[test]
fn array_sum_float_precision() {
    assert_eq!(
        run_prints(r#"<?php echo round(array_sum([0.1, 0.2, 0.3]), 1); "#),
        vec!["0.6"]
    );
}
#[test]
fn compact_returns_array() {
    assert_eq!(
        run_prints(r#"<?php $x=1;$y=2; $r=compact('x','y'); echo gettype($r); "#),
        vec!["array"]
    );
}
#[test]
fn array_diff_returns_missing() {
    assert_eq!(
        run_prints(r#"<?php echo implode(',', array_values(array_diff([3,1,2],[1,2]))); "#),
        vec!["3"]
    );
}

// ── Math ──────────────────────────────────────────────────────

#[test]
fn abs_on_zero() {
    assert_eq!(run_prints(r#"<?php echo abs(0); "#), vec!["0"]);
}
#[test]
fn fmod_exact_zero() {
    assert_eq!(run_prints(r#"<?php echo fmod(6.0, 3.0); "#), vec!["0"]);
}
#[test]
fn pow_fraction_exponent() {
    assert_eq!(run_prints(r#"<?php echo pow(8, 1/3); "#), vec!["2"]);
}
#[test]
fn floor_negative() {
    assert_eq!(run_prints(r#"<?php echo floor(-2.1); "#), vec!["-3"]);
}
#[test]
fn ceil_negative() {
    assert_eq!(run_prints(r#"<?php echo ceil(-2.9); "#), vec!["-2"]);
}
#[test]
fn log_base_self() {
    assert_eq!(run_prints(r#"<?php echo round(log(100, 10)); "#), vec!["2"]);
}
#[test]
fn pi_via_function() {
    assert_eq!(run_prints(r#"<?php echo round(pi(), 4); "#), vec!["3.1416"]);
}

// ── OOP ───────────────────────────────────────────────────────

#[test]
fn interface_constant_in_class() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Codes { const OK = 200; const NOT_FOUND = 404; }
class Response implements Codes {}
echo Response::OK . ',' . Response::NOT_FOUND;
"#
        ),
        vec!["200,404"]
    );
}
#[test]
fn abstract_method_multiple_children() {
    assert_eq!(
        run_prints(
            r#"<?php
abstract class Op { abstract public function run(int $a, int $b): int; }
class Add extends Op { public function run(int $a, int $b): int { return $a+$b; } }
class Mul extends Op { public function run(int $a, int $b): int { return $a*$b; } }
echo (new Add)->run(3,4) . ',' . (new Mul)->run(3,4);
"#
        ),
        vec!["7,12"]
    );
}
#[test]
fn clone_with_array_property() {
    assert_eq!(
        run_prints(
            r#"<?php
class Box { public array $items = []; }
$a = new Box; $a->items[] = 1;
$b = clone $a; $b->items[] = 2;
echo count($a->items) . ',' . count($b->items);
"#
        ),
        vec!["1,2"]
    );
}
#[test]
fn static_property_increments() {
    assert_eq!(
        run_prints(
            r#"<?php
class Counter { static int $n=0; public function __construct(){self::$n++;} }
new Counter; new Counter; new Counter;
echo Counter::$n;
"#
        ),
        vec!["3"]
    );
}
#[test]
fn method_chaining_returns_this() {
    assert_eq!(
        run_prints(
            r#"<?php
class Str { private string $s=''; public function append(string $v):static{$this->s.=$v;return $this;} public function get():string{return $this->s;} }
echo (new Str)->append('a')->append('b')->append('c')->get();
"#
        ),
        vec!["abc"]
    );
}

// ── Control flow ──────────────────────────────────────────────

#[test]
fn while_with_break() {
    assert_eq!(
        run_prints(
            r#"<?php
$i=0; while(true){if($i>=3)break;$i++;} echo $i;
"#
        ),
        vec!["3"]
    );
}
#[test]
fn for_reverse() {
    assert_eq!(
        run_prints(
            r#"<?php
for($i=5;$i>0;$i--) echo $i;
"#
        ),
        vec!["54321"]
    );
}
#[test]
fn switch_default_only() {
    assert_eq!(
        run_prints(r#"<?php switch(99){default:echo 'def';}  "#),
        vec!["def"]
    );
}
#[test]
fn ternary_nested_short() {
    assert_eq!(
        run_prints(r#"<?php $x=5; echo $x>10?'big':($x>3?'mid':'small'); "#),
        vec!["mid"]
    );
}
#[test]
fn match_no_default_throws() {
    assert_eq!(
        run_prints(
            r#"<?php try{$r=match(5){1=>'one',2=>'two'};}catch(\UnhandledMatchError $e){echo 'err';} "#
        ),
        vec!["err"]
    );
}

// ── Type juggling ─────────────────────────────────────────────

#[test]
fn string_false_is_truthy() {
    assert_eq!(
        run_prints(r#"<?php echo (bool)'false' ? 'truthy' : 'falsy'; "#),
        vec!["truthy"]
    );
}
#[test]
fn array_coercion_to_bool() {
    assert_eq!(
        run_prints(r#"<?php echo (bool)[0] ? 'truthy' : 'falsy'; "#),
        vec!["truthy"]
    );
}
#[test]
fn null_coerced_to_zero_in_math() {
    assert_eq!(run_prints(r#"<?php echo null + 10; "#), vec!["10"]);
}
#[test]
fn int_to_bool_in_if() {
    assert_eq!(
        run_prints(r#"<?php if(-1) echo 'yes'; else echo 'no'; "#),
        vec!["yes"]
    );
}
#[test]
fn float_zero_is_falsy() {
    assert_eq!(
        run_prints(r#"<?php echo (bool)0.0 ? 'truthy' : 'falsy'; "#),
        vec!["falsy"]
    );
}

// ── Generators ────────────────────────────────────────────────

#[test]
fn generator_key_is_auto_incremented() {
    assert_eq!(
        run_prints(
            r#"<?php
function gen():Generator{yield 'a';yield 'b';yield 'c';}
foreach(gen() as $k=>$v) echo $k.$v;
"#
        ),
        vec!["0a1b2c"]
    );
}
#[test]
fn generator_send_null_on_first_call() {
    assert_eq!(
        run_prints(
            r#"<?php
function gen():Generator{$v=yield 1;echo $v===null?'null':'notnull';}
$g=gen(); $g->current(); $g->next();
"#
        ),
        vec!["null"]
    );
}

// ── Closures ─────────────────────────────────────────────────

#[test]
fn arrow_fn_ignores_outer_modification() {
    assert_eq!(
        run_prints(r#"<?php $x=1; $f=fn()=>$x; $x=99; echo $f(); "#),
        vec!["1"]
    );
}
#[test]
fn closure_returning_closure() {
    assert_eq!(
        run_prints(
            r#"<?php
$adder=fn($a)=>fn($b)=>$a+$b;
echo $adder(3)(4);
"#
        ),
        vec!["7"]
    );
}
#[test]
fn is_callable_closure() {
    assert_eq!(
        run_prints(r#"<?php $f=fn()=>1; echo is_callable($f)?'yes':'no'; "#),
        vec!["yes"]
    );
}

// ── Misc ──────────────────────────────────────────────────────

#[test]
fn printf_returns_int() {
    assert_eq!(
        run_prints(r#"<?php $n=printf('%s','hi'); echo ' '.$n; "#),
        vec!["hi 2"]
    );
}
#[test]
fn gettype_object() {
    assert_eq!(
        run_prints(r#"<?php echo gettype(new stdClass); "#),
        vec!["object"]
    );
}
#[test]
fn is_object() {
    assert_eq!(
        run_prints(r#"<?php echo is_object(new stdClass)?'yes':'no'; "#),
        vec!["yes"]
    );
}
#[test]
fn get_class_stdclass() {
    assert_eq!(
        run_prints(r#"<?php echo get_class(new stdClass); "#),
        vec!["stdClass"]
    );
}
#[test]
fn property_exists() {
    assert_eq!(
        run_prints(r#"<?php class A{public int $x=1;} echo property_exists('A','x')?'yes':'no'; "#),
        vec!["yes"]
    );
}
#[test]
fn method_exists() {
    assert_eq!(
        run_prints(
            r#"<?php class A{public function foo(){}} echo method_exists('A','foo')?'yes':'no'; "#
        ),
        vec!["yes"]
    );
}
#[test]
fn class_exists() {
    assert_eq!(
        run_prints(
            r#"<?php class MyUniqueClass{} echo class_exists('MyUniqueClass')?'yes':'no'; "#
        ),
        vec!["yes"]
    );
}
#[test]
fn function_exists_builtin() {
    assert_eq!(
        run_prints(r#"<?php echo function_exists('array_map')?'yes':'no'; "#),
        vec!["yes"]
    );
}
#[test]
fn defined_true_false() {
    assert_eq!(
        run_prints(
            r#"<?php define('FOO','bar'); echo defined('FOO')?'yes':'no'; echo defined('BAR_UNDEF')?'yes':'no'; "#
        ),
        vec!["yesno"]
    );
}
#[test]
fn constant_function() {
    assert_eq!(
        run_prints(r#"<?php define('MY_VAL',99); echo constant('MY_VAL'); "#),
        vec!["99"]
    );
}
#[test]
fn php_sapi_name_not_empty() {
    assert_eq!(
        run_prints(r#"<?php echo strlen(php_sapi_name())>0?'ok':'fail'; "#),
        vec!["ok"]
    );
}
