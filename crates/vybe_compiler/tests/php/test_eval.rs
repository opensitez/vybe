//! `eval()` runtime outcomes — return values, scope, and caught runtime errors.
//! Parse-error `eval` cases live in `test_error_handling_deep.rs` / `test_error_handler_distinct_output.rs`.

crate::php_cases! {
    eval_return_statement_yields_value_to_caller => {
        r#"<?php
$v = eval('return 3 + 4;');
echo $v;
"#,
        ["7"]
    };

    eval_without_return_yields_null => {
        r#"<?php
$v = eval('1 + 2;');
echo $v === null ? 'null' : 'set';
"#,
        ["null"]
    };

    eval_can_assign_outer_variable => {
        r#"<?php
eval('$inner = 11;');
echo $inner;
"#,
        ["11"]
    };

    eval_runtime_division_by_zero_is_catchable => {
        r#"<?php
try { eval('return intdiv(9, 0);'); echo 'ok'; }
catch (DivisionByZeroError $e) { echo 'div'; }
"#,
        ["div"]
    };

    eval_throw_inside_string_is_catchable => {
        r#"<?php
try { eval('throw new RuntimeException("from-eval");'); echo 'ok'; }
catch (RuntimeException $e) { echo $e->getMessage(); }
"#,
        ["from-eval"]
    };

    eval_defines_function_callable_afterward => {
        r#"<?php
eval('function eval_add(int $a, int $b): int { return $a + $b; }');
echo eval_add(2, 5);
"#,
        ["7"]
    };

    eval_nested_eval_return_propagates => {
        r#"<?php
$v = eval('return eval("return 6;");');
echo $v;
"#,
        ["6"]
    };

    eval_class_definition_then_instantiate => {
        r#"<?php
eval('class EvalBox { public function tag(): string { return "box"; } }');
echo (new EvalBox())->tag();
"#,
        ["box"]
    };

    eval_foreach_builds_concatenated_output => {
        r#"<?php
$acc = '';
eval('$acc = ""; foreach ([1,2,3] as $n) { $acc .= $n; }');
echo $acc;
"#,
        ["123"]
    };

    eval_isset_on_dynamic_variable_name => {
        r#"<?php
$who = 'name';
eval('$name = "ada";');
echo isset($$who) ? $$who : 'missing';
"#,
        ["ada"]
    };

    eval_unset_clears_variable_from_outer_scope => {
        r#"<?php
$gone = 1;
eval('unset($gone);');
echo isset($gone) ? 'set' : 'unset';
"#,
        ["unset"]
    };

    eval_if_branch_runs_selected_arm => {
        r#"<?php
$flag = true;
eval('if ($flag) { $pick = "yes"; } else { $pick = "no"; }');
echo $pick;
"#,
        ["yes"]
    };

    eval_switch_picks_matching_case => {
        r#"<?php
$code = 2;
eval('switch ($code) { case 2: $out = "two"; break; default: $out = "other"; }');
echo $out;
"#,
        ["two"]
    };

    eval_try_catch_inside_sets_recovery_flag => {
        r#"<?php
eval('try { throw new LogicException("x"); } catch (LogicException $e) { $ok = "caught"; }');
echo $ok;
"#,
        ["caught"]
    };

    eval_match_expression_returns_value => {
        r#"<?php
eval('$m = match (3) { 1 => "a", 3 => "c", default => "z" };');
echo $m;
"#,
        ["c"]
    };

    eval_concatenated_code_string_mutates_array => {
        r#"<?php
$arr = [1];
eval('$arr[] = 2; $arr[] = 3;');
echo implode('-', $arr);
"#,
        ["1-2-3"]
    };

    eval_readonly_result_used_in_arithmetic => {
        r#"<?php
$n = eval('return (int)("12" + 1);');
echo $n;
"#,
        ["13"]
    };

    eval_closing_tag_not_required_in_string => {
        r#"<?php
$v = eval('return strlen("hi");');
echo $v;
"#,
        ["2"]
    };

    eval_type_error_from_strlen_on_array_inside_eval => {
        r#"<?php
try { eval('strlen([]);'); echo 'ok'; }
catch (TypeError $e) { echo 'typed'; }
"#,
        ["typed"]
    };

    eval_generator_function_yields_values => {
        r#"<?php
eval('function eg(): Generator { yield "a"; yield "b"; }');
echo implode('', iterator_to_array(eg()));
"#,
        ["ab"]
    };
}
