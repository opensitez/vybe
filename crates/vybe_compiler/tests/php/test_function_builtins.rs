use super::helpers::compile_ok;

// ── func_get_args / func_num_args / func_get_arg ─────────────────

#[test] fn func_get_args_variadic() { compile_ok(r#"<?php
function gather() {
    $args = func_get_args();
    echo count($args);
    echo implode(',', $args);
}
gather(1, 2, 3);
"#); }

#[test] fn func_num_args_count() { compile_ok(r#"<?php
function count_args() {
    return func_num_args();
}
echo count_args(10, 20, 30);
echo count_args();
"#); }

#[test] fn func_get_arg_by_index() { compile_ok(r#"<?php
function get_second() {
    return func_get_arg(1);
}
echo get_second('first', 'second', 'third');
"#); }

// ── call_user_func / call_user_func_array ────────────────────────

#[test] fn call_user_func_by_name() { compile_ok(r#"<?php
function greet(string $name): string {
    return 'Hello, ' . $name;
}
echo call_user_func('greet', 'World');
"#); }

#[test] fn call_user_func_closure() { compile_ok(r#"<?php
$double = function(int $n): int { return $n * 2; };
echo call_user_func($double, 21);
"#); }

#[test] fn call_user_func_array_basic() { compile_ok(r#"<?php
function add(int $a, int $b): int { return $a + $b; }
echo call_user_func_array('add', [10, 32]);
"#); }

// ── function_exists ───────────────────────────────────────────────

#[test] fn function_exists_builtin() { compile_ok(r#"<?php
echo function_exists('strlen') ? 'yes' : 'no';
echo function_exists('array_map') ? 'yes' : 'no';
echo function_exists('no_such_function_xyz') ? 'yes' : 'no';
"#); }

#[test] fn function_exists_user_defined() { compile_ok(r#"<?php
echo function_exists('myFunc') ? 'yes' : 'no';
function myFunc() {}
echo function_exists('myFunc') ? 'yes' : 'no';
"#); }

// ── forward_static_call / forward_static_call_array ─────────────

#[test] fn forward_static_call_lsb() { compile_ok(r#"<?php
class Base {
    static function create() {
        return forward_static_call(['static', 'build']);
    }
    static function build() {
        return 'base';
    }
}
class Child extends Base {
    static function build() {
        return 'child';
    }
}
echo Child::create();
"#); }

#[test] fn forward_static_call_array_lsb() { compile_ok(r#"<?php
class Logger {
    static function log() {
        return forward_static_call_array(['static', 'format'], func_get_args());
    }
    static function format(string $msg): string {
        return '[LOG] ' . $msg;
    }
}
echo Logger::log('test message');
"#); }

// ── register_shutdown_function ───────────────────────────────────

#[test] fn register_shutdown_function_basic() { compile_ok(r#"<?php
register_shutdown_function(function() {
    echo 'shutdown';
});
echo 'running';
"#); }

// ── get_defined_functions ─────────────────────────────────────────

#[test] fn get_defined_functions_lists() { compile_ok(r#"<?php
function myCustomFn() {}
$all = get_defined_functions();
echo isset($all['user']) ? 'has user' : 'no user';
echo isset($all['internal']) ? ':has internal' : ':no internal';
echo in_array('mycustomfn', $all['user']) ? ':found' : ':not found';
"#); }

// ── array_map with null callback (zip) ───────────────────────────

#[test] fn array_map_null_callback_zip() { compile_ok(r#"<?php
$a = [1, 2, 3];
$b = ['a', 'b', 'c'];
$zipped = array_map(null, $a, $b);
echo count($zipped);
echo $zipped[0][0] . $zipped[0][1];
"#); }

// ── array_walk with key and extra data ───────────────────────────

#[test] fn array_walk_with_key_and_extra() { compile_ok(r#"<?php
$fruits = ['a' => 'apple', 'b' => 'banana', 'c' => 'cherry'];
array_walk($fruits, function(&$value, $key, $prefix) {
    $value = $prefix . ':' . $key . '=' . $value;
}, 'fruit');
echo implode(',', $fruits);
"#); }

// ── usort with callable [$obj, 'method'] ─────────────────────────

#[test] fn usort_with_method_callable() { compile_ok(r#"<?php
class Comparator {
    public function compare($a, $b): int {
        return $a <=> $b;
    }
}
$cmp = new Comparator();
$arr = [3, 1, 4, 1, 5, 9, 2, 6];
usort($arr, [$cmp, 'compare']);
echo implode(',', $arr);
"#); }

// ── usort with static method ['ClassName', 'method'] ─────────────

#[test] fn usort_with_static_method_callable() { compile_ok(r#"<?php
class Sorter {
    public static function descending($a, $b): int {
        return $b <=> $a;
    }
}
$arr = [3, 1, 4, 1, 5, 9, 2, 6];
usort($arr, ['Sorter', 'descending']);
echo implode(',', $arr);
"#); }
