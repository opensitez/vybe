use super::helpers::{compile_ok, run_prints};

// If / elseif / else
#[test]
fn if_simple() {
    compile_ok("<?php if ($x > 0) { echo 'yes'; }");
}
#[test]
fn if_else() {
    compile_ok("<?php if ($x > 0) { echo 'pos'; } else { echo 'neg'; }");
}
#[test]
fn if_elseif_else() {
    compile_ok(
        "<?php if ($x > 0) { echo 'pos'; } elseif ($x < 0) { echo 'neg'; } else { echo 'zero'; }",
    );
}
#[test]
fn nested_if() {
    compile_ok("<?php if ($a) { if ($b) { echo 'both'; } }");
}
#[test]
fn if_alternative_syntax() {
    compile_ok("<?php if ($x > 0) : echo 'yes'; endif;");
}
#[test]
fn if_alternative_syntax_with_elseif_else() {
    compile_ok(
        "<?php if ($x > 0) : echo 'pos'; elseif ($x < 0) : echo 'neg'; else : echo 'zero'; endif;",
    );
}
#[test]
fn if_alternative_syntax_wraps_polyfill_function() {
    compile_ok(
        r#"<?php
if ( ! function_exists( 'mb_substr' ) ) :
	function mb_substr( $text, $start, $length = null, $encoding = null ) {
		return _mb_substr( $text, $start, $length, $encoding );
	}
endif;
"#,
    );
}

// While
#[test]
fn while_loop() {
    compile_ok("<?php $i = 0; while ($i < 10) { $i++; }");
}
#[test]
fn do_while() {
    compile_ok("<?php $i = 0; do { $i++; } while ($i < 10);");
}

// For
#[test]
fn for_loop() {
    compile_ok("<?php for ($i = 0; $i < 10; $i++) { echo $i; }");
}
#[test]
fn for_no_init() {
    compile_ok("<?php $i = 0; for (; $i < 10; $i++) {}");
}
#[test]
fn for_infinite() {
    compile_ok("<?php for (;;) { break; }");
}

// Foreach
#[test]
fn foreach_value() {
    compile_ok("<?php foreach ([1,2,3] as $v) { echo $v; }");
}
#[test]
fn foreach_key_value() {
    compile_ok("<?php foreach (['a'=>1] as $k => $v) { echo $k . $v; }");
}
#[test]
fn foreach_nested() {
    compile_ok("<?php foreach ([[1,2],[3,4]] as $row) { foreach ($row as $v) { echo $v; } }");
}

// Switch
#[test]
fn switch_basic() {
    compile_ok(
        "<?php switch ($x) { case 1: echo 'one'; break; case 2: echo 'two'; break; default: echo 'other'; }",
    );
}
#[test]
fn switch_fallthrough() {
    compile_ok("<?php switch ($x) { case 1: case 2: echo 'one or two'; break; }");
}

// Break / Continue
#[test]
fn break_in_loop() {
    compile_ok("<?php for ($i=0;$i<10;$i++) { if ($i==5) break; }");
}
#[test]
fn continue_in_loop() {
    compile_ok("<?php for ($i=0;$i<10;$i++) { if ($i==3) continue; echo $i; }");
}

// Match (PHP 8)
#[test]
fn match_expr() {
    compile_ok("<?php $x = match($v) { 1 => 'one', 2 => 'two', default => 'other' };");
}

#[test]
fn conditional_runtime_nested_true_false_paths() {
    let out = run_prints(
        "<?php\n$x = 10;\n$y = 3;\nif ($x > 0 && $y < 5) { echo 'in'; } else { echo 'out'; }\n",
    );
    assert_eq!(out, vec!["in"]);
}

#[test]
fn ternary_and_ternary_short_circuit() {
    let out = run_prints(
        "<?php\n$label = true ? 'yes' : 'no';\n$zero = 0;\n$val = $zero ?: 'fallback';\necho $label;\necho '|';\necho $val;\n",
    );
    assert_eq!(out, vec!["yes|fallback"]);
}

#[test]
fn switch_runtime_multi_case() {
    let out = run_prints(
        "<?php\n$x = 2;\nswitch ($x) {\n    case 1:\n        echo 'one';\n        break;\n    case 2:\n    case 3:\n        echo 'two_or_three';\n        break;\n    default:\n        echo 'other';\n}\n",
    );
    assert_eq!(out, vec!["two_or_three"]);
}

#[test]
fn match_runtime_with_array_key() {
    let out = run_prints(
        "<?php\n$role = 'admin';\n$out = match($role) {\n    'admin' => 'full',\n    'guest' => 'read',\n    default => 'none',\n};\necho $out;\n",
    );
    assert_eq!(out, vec!["full"]);
}

#[test]
fn if_coalesce_and_elvis() {
    let out = run_prints(
        "<?php\n$a = null;\n$b = '';\necho $a ?? 'a';\necho '|';\necho $b ?: 'b';\n",
    );
    assert_eq!(out, vec!["a|b"]);
}

#[test]
fn guard_continue_and_break_in_nested_loops() {
    let out = run_prints(
        "<?php\n$total = 0;\nfor ($i = 0; $i < 3; $i++) {\n    foreach ([1, 2, 3] as $j) {\n        if ($j === 2) { continue; }\n        if ($j === 3 && $i === 1) { break 2; }\n        $total += $j;\n    }\n}\necho $total;\n",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn for_loop_multiple_update_expressions_runtime() {
    let out = run_prints(
        "<?php\n$total = 0;\nfor ($i = 0, $j = 0; $i < 4; $i++, $j++) {\n    if ($i === 2) {\n        continue;\n    }\n    $total += $i + $j;\n}\necho $total;\n",
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn while_loop_continue_and_break_runtime() {
    let out = run_prints(
        "<?php\n$i = 0;\n$j = 0;\nwhile ($i < 6) {\n    $i++;\n    if ($i === 2 || $i === 3) {\n        continue;\n    }\n    if ($i === 5) {\n        break;\n    }\n    $j += $i;\n}\necho $j;\n",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn do_while_executes_once_with_zero_condition_runtime() {
    let out = run_prints(
        "<?php\n$i = 0;\ndo {\n    echo $i;\n    $i++;\n} while (false);\n",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn foreach_by_reference_mutates_source_runtime() {
    let out = run_prints(
        "<?php\n$items = [1, 2, 3];\nforeach ($items as &$value) {\n    $value *= 2;\n}\nunset($value);\necho $items[0];\necho ',';\necho $items[1];\necho ',';\necho $items[2];\n",
    );
    assert_eq!(out, vec!["2,4,6"]);
}

#[test]
fn switch_fallthrough_without_explicit_break_runtime() {
    let out = run_prints(
        "<?php\n$kind = 'warning';\n$state = '';\nswitch ($kind) {\n    case 'error':\n        $state .= 'error';\n        // intentional fallthrough\n    case 'warning':\n        $state .= 'warning';\n        // intentional fallthrough\n    case 'notice':\n        $state .= 'notice';\n        break;\n    default:\n        $state .= 'unknown';\n}\necho $state;\n",
    );
    assert_eq!(out, vec!["warningnotice"]);
}

#[test]
fn match_with_default_and_expression_runtime() {
    let out = run_prints(
        "<?php\n$value = 9;\n$out = match (true) {\n    $value > 10 => 'gt10',\n    $value > 5 => 'gt5',\n    default => 'other',\n};\necho $out;\n",
    );
    assert_eq!(out, vec!["gt5"]);
}

#[test]
fn if_with_match_in_condition_runtime() {
    let out = run_prints(
        "<?php\n$code = 2;\nif (match ($code) { 1 => false, 2 => true, default => false }) {\n    echo 'on';\n} else {\n    echo 'off';\n}\n",
    );
    assert_eq!(out, vec!["on"]);
}

#[test]
fn nested_if_and_match_subject_equality_runtime() {
    let out = run_prints(
        "<?php\n$level = 3;\n$state = match (true) {\n    $level > 2 && $level < 10 => 'inner',\n    default => 'outer',\n};\nif ($state === 'inner') {\n    echo 'in';\n} else {\n    echo 'out';\n}\n",
    );
    assert_eq!(out, vec!["in"]);
}

#[test]
fn conditional_with_assignment_precedence_in_condition_runtime() {
    let out = run_prints(
        "<?php\n$flag = 0;\nif (($flag = 3) > 1 && $flag === 3) {\n    echo $flag . '|';\n}\nif ($flag === 3) {\n    echo 'done';\n}\n",
    );
    assert_eq!(out, vec!["3|done"]);
}

#[test]
fn if_condition_with_nested_logical_groups_runtime() {
    let out = run_prints(
        "<?php\n$a = true;\n$b = false;\n$c = true;\nif (($a && $b) || (!$b && $c)) {\n    echo 'hit';\n} else {\n    echo 'miss';\n}\n",
    );
    assert_eq!(out, vec!["hit"]);
}

#[test]
fn if_uses_assignment_in_nested_condition_runtime() {
    let out = run_prints(
        "<?php\n$a = 0;\nif (($a = 2) > 1 && $a < 4) {\n    echo $a;\n}\n",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn switch_default_last_case_with_fallthrough_runtime() {
    let out = run_prints(
        "<?php\n$mode = 'x';\n$out = '';\nswitch ($mode) {\n    case 'a':\n        $out = 'a';\n        break;\n    default:\n        $out = 'd';\n    case 'b':\n        $out .= '|b';\n        break;\n}\necho $out;\n",
    );
    assert_eq!(out, vec!["d|b"]);
}

#[test]
fn break_and_continue_with_level_1_inside_nested_loops_runtime() {
    let out = run_prints(
        "<?php\n$total = 0;\nfor ($i = 0; $i < 3; $i++) {\n    for ($j = 0; $j < 3; $j++) {\n        if ($j === 1) {\n            continue 1;\n        }\n        if ($j === 2) {\n            break 1;\n        }\n        $total += 1;\n    }\n}\necho $total;\n",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn do_while_condition_update_by_reference_like_runtime() {
    let out = run_prints(
        "<?php\n$i = 0;\n$count = 0;\ndo {\n    $count++;\n    $i++;\n    if ($i === 2) {\n        continue;\n    }\n} while ($i < 4);\necho $count;\n",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn break_and_continue_with_nested_level_runtime() {
    let out = run_prints(
        "<?php\n$total = 0;\nfor ($i = 0; $i < 3; $i++) {\n    for ($j = 0; $j < 4; $j++) {\n        if ($j === 1) {\n            continue 2;\n        }\n        if ($j === 3) {\n            break 2;\n        }\n        $total++;\n    }\n}\necho $total;\n",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn switch_with_expression_case_runtime() {
    let out = run_prints(
        "<?php\n$mode = 4;\n$out = '';\nswitch ($mode) {\n    case 1 + 1:\n        $out .= '2';\n        break;\n    case 2 + 2:\n        $out .= '4';\n        break;\n    default:\n        $out .= 'd';\n}\necho $out;\n",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn if_elseif_chain_truthy_falsy_runtime() {
    let out = run_prints(
        "<?php\n$x = '';\nif ($x) {\n    echo 'x';\n} elseif ('0') {\n    echo 'zero-string';\n} else {\n    echo 'else';\n}\n",
    );
    assert_eq!(out, vec!["zero-string"]);
}

#[test]
fn while_loop_with_complex_condition_runtime() {
    let out = run_prints(
        "<?php\n$i = 0;\n$j = 0;\nwhile (($i < 4) && ($j < 2)) {\n    $i++;\n    if ($i % 2 === 0) {\n        continue;\n    }\n    $j++;\n}\necho $i . '|' . $j;\n",
    );
    assert_eq!(out, vec!["3|2"]);
}

#[test]
fn if_nested_assignment_in_else_if_runtime() {
    let out = run_prints(
        "<?php\n$state = 'init';\nif (false) {\n    echo 'bad';\n} elseif (($state = 'ready') && true) {\n    echo $state;\n} else {\n    echo 'final';\n}\n",
    );
    assert_eq!(out, vec!["ready"]);
}
