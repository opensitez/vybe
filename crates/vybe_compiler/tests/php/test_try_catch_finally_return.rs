//! Finally block return/throw overriding try/catch; break/continue; nested order.

crate::php_cases! {
    finally_return_overrides_try_value => {
        r#"<?php
function f(): string {
    try { return 'try'; }
    finally { return 'finally'; }
}
echo f();
"#,
        ["finally"]
    };

    finally_return_overrides_catch_value => {
        r#"<?php
function f(): string {
    try { throw new Exception('x'); }
    catch (Exception $e) { return 'catch'; }
    finally { return 'finally'; }
}
echo f();
"#,
        ["finally"]
    };

    finally_throw_overrides_try_return => {
        r#"<?php
function f(): void {
    try { return; }
    finally { throw new Exception('from finally'); }
}
try { f(); echo 'after'; }
catch (Exception $e) { echo $e->getMessage(); }
"#,
        ["from finally"]
    };

    empty_finally_preserves_try_return => {
        r#"<?php
function f(): int {
    try { return 42; }
    finally {}
}
echo f();
"#,
        ["42"]
    };

    finally_side_effects_before_try_return => {
        r#"<?php
function f(): string {
    try { return 'ok'; }
    finally { echo 'cleanup,'; }
}
echo f();
"#,
        ["cleanup,ok"]
    };

    catch_return_then_finally_overrides => {
        r#"<?php
function f(): string {
    try { throw new RuntimeException('boom'); }
    catch (RuntimeException $e) { return 'caught'; }
    finally { return 'done'; }
}
echo f();
"#,
        ["done"]
    };

    nested_finally_inner_return_wins => {
        r#"<?php
function f(): string {
    try {
        try { return 'inner try'; }
        finally { return 'inner finally'; }
    }
    finally { return 'outer finally'; }
}
echo f();
"#,
        ["inner finally"]
    };

    nested_finally_outer_return_when_inner_no_return => {
        r#"<?php
function f(): string {
    try {
        try { echo 'work,'; }
        finally { echo 'inner,'; }
    }
    finally { return 'outer'; }
}
echo f();
"#,
        ["work,inner,outer"]
    };

    finally_runs_on_foreach_break => {
        r#"<?php
$log = [];
foreach ([1, 2, 3] as $n) {
    try {
        $log[] = "t$n";
        if ($n === 2) { break; }
    } finally {
        $log[] = "f$n";
    }
}
echo implode(',', $log);
"#,
        ["t1,f1,t2,f2"]
    };

    finally_runs_on_foreach_continue => {
        r#"<?php
$log = [];
foreach ([1, 2, 3] as $n) {
    try {
        $log[] = "t$n";
        if ($n === 2) { continue; }
        $log[] = "a$n";
    } finally {
        $log[] = "f$n";
    }
}
echo implode(',', $log);
"#,
        ["t1,a1,f1,t2,f2,t3,a3,f3"]
    };

    finally_break_exits_while_loop => {
        r#"<?php
$log = [];
$n = 0;
while ($n < 5) {
    try {
        $log[] = $n;
        if ($n === 2) { throw new Exception('stop'); }
    } catch (Exception $e) {
        $log[] = 'c';
    } finally {
        if ($n === 2) { break; }
    }
    $n++;
}
echo implode(',', $log);
"#,
        ["0,1,2,c"]
    };

    finally_continue_skips_while_iteration => {
        r#"<?php
$log = [];
$n = 0;
while ($n < 4) {
    try {
        $log[] = "t$n";
        if ($n === 1) { throw new Exception('skip'); }
        $log[] = "a$n";
    } catch (Exception $e) {
        $log[] = 'c';
    } finally {
        if ($n === 1) { $n++; continue; }
    }
    $n++;
}
echo implode(',', $log);
"#,
        ["t0,a0,t1,c,t2,a2,t3,a3"]
    };

    finally_break_from_labeled_loop => {
        r#"<?php
$log = [];
outer: for ($i = 0; $i < 4; $i++) {
    try {
        $log[] = $i;
        if ($i === 2) { throw new Exception('x'); }
    } finally {
        if ($i === 2) { break outer; }
    }
}
echo implode(',', $log);
"#,
        ["0,1,2"]
    };

    finally_continue_in_for_loop => {
        r#"<?php
$log = [];
for ($i = 0; $i < 4; $i++) {
    try {
        if ($i === 1) { throw new Exception('cont'); }
        $log[] = "ok$i";
    } finally {
        if ($i === 1) { continue; }
    }
}
echo implode(',', $log);
"#,
        ["ok0,ok2,ok3"]
    };

    finally_runs_after_catch_handles_throw => {
        r#"<?php
$log = [];
try {
    throw new Exception('e');
} catch (Exception $ex) {
    $log[] = 'catch';
} finally {
    $log[] = 'finally';
}
echo implode(',', $log);
"#,
        ["catch,finally"]
    };

    finally_runs_when_try_throws_to_outer_catch => {
        r#"<?php
$log = [];
try {
    try { throw new Exception('inner'); }
    finally { $log[] = 'inner_finally'; }
} catch (Exception $e) {
    $log[] = 'outer_catch';
}
echo implode(',', $log);
"#,
        ["inner_finally,outer_catch"]
    };

    throw_from_finally_replaces_try_exception => {
        r#"<?php
try {
    try { throw new Exception('try'); }
    finally { throw new Exception('finally'); }
} catch (Exception $e) {
    echo $e->getMessage();
}
"#,
        ["finally"]
    };

    throw_from_finally_supersedes_catch_return => {
        r#"<?php
function f(): string {
    try { throw new Exception('t'); }
    catch (Exception $e) { return 'catch'; }
    finally { throw new Exception('finally throw'); }
}
try { echo f(); }
catch (Exception $e) { echo $e->getMessage(); }
"#,
        ["finally throw"]
    };

    return_from_catch_blocked_by_finally_return => {
        r#"<?php
function f(): string {
    try { throw new InvalidArgumentException('bad'); }
    catch (InvalidArgumentException $e) { return 'handled'; }
    finally { return 'final'; }
}
echo f();
"#,
        ["final"]
    };

    finally_return_in_closure => {
        r#"<?php
$fn = function (): string {
    try { return 'closure try'; }
    finally { return 'closure finally'; }
};
echo $fn();
"#,
        ["closure finally"]
    };

    finally_return_in_class_method => {
        r#"<?php
class Worker {
    public function run(): string {
        try { return 'method try'; }
        finally { return 'method finally'; }
    }
}
echo (new Worker())->run();
"#,
        ["method finally"]
    };

    finally_return_null_overrides_try_int => {
        r#"<?php
function f(): ?int {
    try { return 99; }
    finally { return null; }
}
var_export(f());
"#,
        ["NULL"]
    };

    finally_return_string_overrides_try_int => {
        r#"<?php
function f(): string {
    try { return (string) 7; }
    finally { return 'text'; }
}
echo f();
"#,
        ["text"]
    };

    finally_echo_only_does_not_override_return => {
        r#"<?php
function f(): int {
    try { return 5; }
    finally { echo 'log,'; }
}
echo f();
"#,
        ["log,5"]
    };

    finally_return_in_triple_nested_try => {
        r#"<?php
function f(): string {
    try {
        try {
            try { return 'deep'; }
            finally { return 'level3'; }
        } finally { echo 'l2,'; }
    } finally { return 'level1'; }
}
echo f();
"#,
        ["l2,level3"]
    };

    inner_finally_runs_before_outer_finally_return => {
        r#"<?php
$log = [];
function f() use (&$log): string {
    try {
        try { $log[] = 'try'; return 'inner'; }
        finally { $log[] = 'inner_f'; return 'inner_ret'; }
    } finally {
        $log[] = 'outer_f';
        return 'outer_ret';
    }
}
echo f() . ':' . implode(',', $log);
"#,
        ["inner_ret:try,inner_f,outer_f"]
    };

    finally_runs_on_switch_break => {
        r#"<?php
$log = [];
switch (2) {
    case 2:
        try {
            $log[] = 'case';
            break;
        } finally {
            $log[] = 'fin';
        }
}
echo implode(',', $log);
"#,
        ["case,fin"]
    };

    finally_runs_each_loop_iteration => {
        r#"<?php
$log = [];
for ($i = 0; $i < 3; $i++) {
    try { $log[] = "b$i"; }
    finally { $log[] = "e$i"; }
}
echo implode(',', $log);
"#,
        ["b0,e0,b1,e1,b2,e2"]
    };

    try_return_finally_modifies_outer_var => {
        r#"<?php
$state = 'init';
function work() use (&$state): string {
    try { return 'ret'; }
    finally { $state = 'cleaned'; }
}
$v = work();
echo $state . ':' . $v;
"#,
        ["cleaned:ret"]
    };

    finally_return_in_anonymous_function => {
        r#"<?php
$fn = function (): string {
    try { throw new Exception('x'); }
    catch (Exception $e) { return 'no'; }
    finally { return 'yes'; }
};
echo $fn();
"#,
        ["yes"]
    };

    finally_break_exits_do_while => {
        r#"<?php
$log = [];
$n = 0;
do {
    try {
        $log[] = $n;
        if ($n === 1) { break; }
    } finally {
        $log[] = "f$n";
    }
    $n++;
} while ($n < 5);
echo implode(',', $log);
"#,
        ["0,f0,1,f1"]
    };

    finally_continue_do_while_loop => {
        r#"<?php
$log = [];
$n = 0;
do {
    try {
        $log[] = "t$n";
        if ($n === 0) { throw new Exception('skip'); }
    } catch (Exception $e) {
        $log[] = 'c';
    } finally {
        if ($n === 0) { $n++; continue; }
    }
    $n++;
} while ($n < 3);
echo implode(',', $log);
"#,
        ["t0,c,t1,t2"]
    };

    nested_try_finally_return_chain_innermost_wins => {
        r#"<?php
function f(): int {
    try {
        try { return 1; }
        finally { return 2; }
    } finally {
        return 3;
    }
}
echo f();
"#,
        ["2"]
    };

    finally_throw_caught_by_outer_catch => {
        r#"<?php
try {
    try { return; }
    finally { throw new RuntimeException('outer path'); }
} catch (RuntimeException $e) {
    echo $e->getMessage();
}
"#,
        ["outer path"]
    };

    finally_return_after_catch_rethrow_attempt => {
        r#"<?php
function f(): string {
    try { throw new Exception('orig'); }
    catch (Exception $e) { throw $e; }
    finally { return 'stopped'; }
}
try { echo f(); }
catch (Exception $e) { echo 'leaked'; }
"#,
        ["stopped"]
    };

    finally_empty_with_try_return_in_function => {
        r#"<?php
function answer(): int {
    try { return 100; }
    finally {}
}
echo answer();
"#,
        ["100"]
    };

    finally_return_overrides_return_after_assignment_in_try => {
        r#"<?php
function f(): string {
    try {
        $x = 'computed';
        return $x;
    } finally {
        return 'override';
    }
}
echo f();
"#,
        ["override"]
    };

    finally_runs_before_return_value_reaches_caller => {
        r#"<?php
$order = [];
function f() use (&$order): string {
    try {
        $order[] = 'try';
        return 'value';
    } finally {
        $order[] = 'finally';
    }
}
echo f() . ':' . implode(',', $order);
"#,
        ["value:try,finally"]
    };

    finally_return_with_typed_string_return => {
        r#"<?php
function label(): string {
    try { return 'alpha'; }
    finally { return 'beta'; }
}
echo label();
"#,
        ["beta"]
    };

    finally_return_inside_trait_method => {
        r#"<?php
trait Service {
    public function run(): string {
        try { return 'try'; }
        finally { return 'finally'; }
    }
}
class App { use Service; }
echo (new App())->run();
"#,
        ["finally"]
    };

    finally_return_in_static_method => {
        r#"<?php
class Calc {
    public static function compute(): string {
        try { return 'static try'; }
        finally { return 'static finally'; }
    }
}
echo Calc::compute();
"#,
        ["static finally"]
    };

    try_catch_finally_all_return_finally_wins => {
        r#"<?php
function f(): string {
    try { return 'try'; }
    catch (Exception $e) { return 'catch'; }
    finally { return 'finally'; }
}
echo f();
"#,
        ["finally"]
    };

    finally_throw_overrides_catch_throw => {
        r#"<?php
try {
    try { throw new Exception('first'); }
    catch (Exception $e) { throw new Exception('second'); }
    finally { throw new Exception('third'); }
} catch (Exception $e) {
    echo $e->getMessage();
}
"#,
        ["third"]
    };

    finally_return_after_successful_try_no_throw => {
        r#"<?php
function ok(): string {
    try {
        echo 'prep,';
        return 'result';
    } finally {
        return 'final';
    }
}
echo ok();
"#,
        ["prep,final"]
    };

    finally_in_try_with_outer_catch_on_finally_throw => {
        r#"<?php
$log = [];
try {
    try { $log[] = 'inner'; }
    finally {
        $log[] = 'fin';
        throw new LogicException('logic');
    }
} catch (LogicException $e) {
    $log[] = 'caught';
}
echo implode(',', $log);
"#,
        ["inner,fin,caught"]
    };

    finally_break_only_exits_innermost_loop => {
        r#"<?php
$log = [];
for ($i = 0; $i < 2; $i++) {
    for ($j = 0; $j < 3; $j++) {
        try {
            $log[] = "$i$j";
            if ($j === 1) { throw new Exception('b'); }
        } finally {
            if ($j === 1) { break; }
        }
    }
}
echo implode(',', $log);
"#,
        ["00,01"]
    };

    finally_return_overrides_early_return_in_if => {
        r#"<?php
function pick(bool $flag): string {
    try {
        if ($flag) { return 'early'; }
        return 'late';
    } finally {
        return 'always';
    }
}
echo pick(true) . pick(false);
"#,
        ["alwaysalways"]
    };

    nested_finally_throw_reaches_outer_catch => {
        r#"<?php
try {
    try { echo 'start,'; }
    finally { throw new DomainException('domain'); }
} catch (DomainException $e) {
    echo $e->getMessage();
}
"#,
        ["start,domain"]
    };

    finally_logs_on_break_without_return_override => {
        r#"<?php
$log = [];
foreach (['a', 'b'] as $ch) {
    try {
        $log[] = $ch;
        break;
    } finally {
        $log[] = 'fin';
    }
}
echo implode(',', $log);
"#,
        ["a,fin"]
    };

    finally_with_unset_does_not_change_return => {
        r#"<?php
function f(): int {
    $tmp = 9;
    try { return $tmp; }
    finally { unset($tmp); }
}
echo f();
"#,
        ["9"]
    };

    finally_return_overrides_try_in_switch_case => {
        r#"<?php
function run(int $v): string {
    switch ($v) {
        case 1:
            try { return 'case try'; }
            finally { return 'case finally'; }
    }
    return 'miss';
}
echo run(1);
"#,
        ["case finally"]
    };

    finally_return_propagates_from_called_function => {
        r#"<?php
function inner(): string {
    try { return 'inner'; }
    finally { return 'from inner'; }
}
function outer(): string {
    return inner();
}
echo outer();
"#,
        ["from inner"]
    };

    finally_return_after_try_with_multiple_echo => {
        r#"<?php
function f(): string {
    try {
        echo 'a,';
        echo 'b,';
        return 'c';
    } finally {
        return 'd';
    }
}
echo f();
"#,
        ["a,b,d"]
    };

    finally_runs_when_catch_returns_without_finally_return => {
        r#"<?php
function f(): string {
    try { throw new Exception('x'); }
    catch (Exception $e) { echo 'handled,'; return 'ok'; }
    finally { echo 'cleanup,'; }
}
echo f();
"#,
        ["handled,cleanup,ok"]
    };

    finally_return_overrides_void_try_early_exit => {
        r#"<?php
function f(): string {
    try {
        if (true) { return 'try exit'; }
    } finally {
        return 'finally exit';
    }
    return 'unreachable';
}
echo f();
"#,
        ["finally exit"]
    };

    finally_continue_after_catch_in_foreach => {
        r#"<?php
$log = [];
foreach ([1, 2] as $n) {
    try {
        if ($n === 1) { throw new Exception('e'); }
        $log[] = "ok$n";
    } catch (Exception $ex) {
        $log[] = 'c';
    } finally {
        if ($n === 1) { continue; }
    }
}
echo implode(',', $log);
"#,
        ["c,ok2"]
    };

    finally_return_inherited_method => {
        r#"<?php
class Base {
    public function go(): string {
        try { return 'base'; }
        finally { return 'base finally'; }
    }
}
class Child extends Base {}
echo (new Child())->go();
"#,
        ["base finally"]
    };
}
