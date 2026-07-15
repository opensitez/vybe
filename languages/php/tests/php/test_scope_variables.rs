//! Variable scope: global, static, `use`, and `$GLOBALS`.

use super::helpers::run_prints_dynamic;
use std::time::{SystemTime, UNIX_EPOCH};

crate::php_cases! {
    global_keyword_imports_global => {
        r#"<?php
$x = 1;
function f(): void { global $x; $x = 2; }
f();
echo $x;
"#,
        ["2"]
    };

    globals_superglobal_read_write => {
        r#"<?php
$GLOBALS['g'] = 5;
echo $GLOBALS['g'];
"#,
        ["5"]
    };

    static_local_in_function => {
        r#"<?php
function tick(): int { static $n = 0; return ++$n; }
echo tick() . tick();
"#,
        ["12"]
    };

    static_in_method => {
        r#"<?php
class C {
    public static function n(): int { static $s = 0; return ++$s; }
}
echo C::n() . C::n();
"#,
        ["12"]
    };

    unset_global_removes_binding => {
        r#"<?php
$z = 1;
unset($z);
echo isset($z) ? 'yes' : 'no';
"#,
        ["no"]
    };

    variable_name_does_not_collide_with_error_class => {
        r#"<?php
echo isset($error) ? 'bad' : 'ok';
$error = 'set';
echo isset($error) ? ':' . $error : ':bad';
"#,
        ["ok:set"]
    };

    branch_assigned_variable_survives_block_scope => {
        r#"<?php
function label($key, $flag = null) {
    if ($flag === null) {
        $retval = $key;
    } else {
        $retval = 'bad';
    }
    return $retval;
}
echo label('Host');
"#,
        ["Host"]
    };

    variable_variables => {
        r#"<?php
$a = 'name';
$name = 'bob';
echo $$a;
"#,
        ["bob"]
    };

    variable_function_call => {
        r#"<?php
function hi(): string { return 'hi'; }
$f = 'hi';
echo $f();
"#,
        ["hi"]
    };

    variable_class_instantiation => {
        r#"<?php
class T { public function __construct(public string $k) {} }
$c = 'T';
echo (new $c('x'))->k;
"#,
        ["x"]
    };

    extract_imports_to_local_scope => {
        r#"<?php
$arr = ['a' => 1, 'b' => 2];
extract($arr);
echo $a + $b;
"#,
        ["3"]
    };

    compact_exports_locals => {
        r#"<?php
$x = 1;
$y = 2;
$c = compact('x', 'y');
echo $c['x'] + $c['y'];
"#,
        ["3"]
    };

    list_destructure_assign => {
        r#"<?php
[$a, $b] = [1, 2];
echo $a . $b;
"#,
        ["12"]
    };

    list_with_keys_in_array_destructure => {
        r#"<?php
['x' => $x] = ['x' => 9];
echo $x;
"#,
        ["9"]
    };

    foreach_by_reference_modifies_source => {
        r#"<?php
$a = [1, 2];
foreach ($a as &$v) { $v *= 10; }
echo implode(',', $a);
"#,
        ["10,20"]
    };

    closure_use_imports_global_not_needed => {
        r#"<?php
$outer = 3;
$fn = function () use ($outer): int { return $outer; };
echo $fn();
"#,
        ["3"]
    };

    nested_function_scope_isolated => {
        r#"<?php
function outer(): int {
    $x = 1;
    $inner = function (): int { $y = 2; return $y; };
    return $inner();
}
echo outer();
"#,
        ["2"]
    };

    require_once_defines_function_once => {
        r#"<?php
function local_fn(): string { return 'ok'; }
echo local_fn();
"#,
        ["ok"]
    };

    constant_in_namespace => {
        r#"<?php
namespace ScopeTest;
const K = 'v';
echo K;
"#,
        ["v"]
    };

    define_runtime_constant => {
        r#"<?php
define('MY_FLAG', true);
echo MY_FLAG ? '1' : '0';
"#,
        ["1"]
    };

    defined_check => {
        r#"<?php
echo defined('PHP_VERSION') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    get_defined_constants_has_php_version => {
        r#"<?php
$c = get_defined_constants();
echo isset($c['PHP_VERSION']) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    static_variable_retains_between_calls => {
        r#"<?php
function acc(int $n): int { static $s = 0; $s += $n; return $s; }
echo acc(1) . acc(2);
"#,
        ["13"]
    };

    reference_assignment_aliases => {
        r#"<?php
$a = 1;
$b = &$a;
$b = 2;
echo $a;
"#,
        ["2"]
    };

    unset_reference_leaves_other => {
        r#"<?php
$a = 1;
$b = &$a;
unset($b);
$a = 3;
echo $a;
"#,
        ["3"]
    };

    global_array_element => {
        r#"<?php
$GLOBALS['items'] = [1];
function add(): void { global $items; $items[] = 2; }
add();
echo count($GLOBALS['items']);
"#,
        ["2"]
    };

    static_class_property_shared => {
        r#"<?php
class S { public static int $n = 0; }
S::$n = 4;
echo S::$n;
"#,
        ["4"]
    };

    instance_property_shadows_nothing => {
        r#"<?php
class P { public int $n = 1; }
$p = new P();
$p->n = 2;
echo $p->n;
"#,
        ["2"]
    };

    foreach_key_value_scope => {
        r#"<?php
$m = ['a' => 1];
foreach ($m as $k => $v) { echo $k . $v; }
"#,
        ["a1"]
    };

    switch_case_no_break_fallthrough => {
        r#"<?php
$n = 1;
switch ($n) { case 1: echo 'a'; case 2: echo 'b'; default: echo 'c'; }
"#,
        ["abc"]
    };

    do_while_runs_at_least_once => {
        r#"<?php
$i = 0;
do { $i++; } while ($i < 1);
echo $i;
"#,
        ["1"]
    };

    while_break_exits_loop => {
        r#"<?php
$i = 0;
while (true) { $i++; if ($i === 2) break; }
echo $i;
"#,
        ["2"]
    };

    for_loop_counter => {
        r#"<?php
$s = 0;
for ($i = 1; $i <= 3; $i++) { $s += $i; }
echo $s;
"#,
        ["6"]
    };

    goto_label_jump => {
        r#"<?php
goto end;
echo 'skip';
end:
echo 'done';
"#,
        ["done"]
    };

    include_variable_scope => {
        r#"<?php
$x = 2;
function getx(): int { global $x; return $x; }
echo getx();
"#,
        ["2"]
    };

    isset_on_nested_array_key => {
        r#"<?php
$a = ['k' => ['inner' => 1]];
echo isset($a['k']['inner']) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    empty_on_unset_variable => {
        r#"<?php
echo empty($never_set) ? 'empty' : 'set';
"#,
        ["empty"]
    };

    null_coalesce_unset_variable => {
        r#"<?php
echo $missing ?? 'def';
"#,
        ["def"]
    };
}

#[test]
fn method_include_returns_to_caller_frame() {
    let stamp = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("clock")
        .as_nanos();
    let base = std::env::temp_dir().join(format!("vybe-php-method-include-{stamp}"));
    std::fs::create_dir_all(&base).expect("create temp dir");
    let view_path = base.join("view.php");
    std::fs::write(&view_path, "<?php echo 'T';").expect("write view");

    let src = format!(
        r#"<?php
class C {{
    public function index() {{
        include '{}';
    }}
}}
$action = 'index';
$id = null;
try {{
    $controller = new C();
    switch ($action) {{
        case 'index':
            $controller->index();
            break;
        case 'view':
            if ($id === null) throw new Exception('Project ID required');
            break;
    }}
    echo 'D';
}} catch (Exception $e) {{
    echo 'caught:' . $e->getMessage();
}}
"#,
        view_path
            .to_string_lossy()
            .replace('\\', "\\\\")
            .replace('\'', "\\'")
    );

    assert_eq!(
        run_prints_dynamic(&src, base.join("main.php").to_string_lossy().as_ref()),
        vec!["TD".to_string()]
    );
}
