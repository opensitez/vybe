use super::helpers::run_prints;

fn assert_int(expr: &str, expected: i64) {
    assert_eq!(
        run_prints(&format!("<?php echo {}; ", expr)),
        vec![expected.to_string()]
    );
}

#[test]
fn php_control_flow_constructs() {
    for n in 1..=5_i64 {
        let for_expected = n;
        let while_expected = n * (n + 1) / 2;
        let do_expected = 2 * n;
        let foreach_expected = n;

        assert_int(
            &format!("$sum = 0; for ($i = 0; $i < {n}; $i++) {{ $sum += 1; }} echo $sum;"),
            for_expected,
        );
        assert_int(
            &format!("$sum = 0; $i = {n}; while ($i > 0) {{ $sum += $i; $i--; }} echo $sum;"),
            while_expected,
        );
        assert_int(
            &format!("$sum = 0; $i = 0; do {{ $sum += 2; $i++; }} while ($i < {n}); echo $sum;"),
            do_expected,
        );
        assert_int(
            &format!(
                "$sum = 0; foreach (array_fill(0, {n}, 1) as $value) {{ $sum += $value; }} echo $sum;"
            ),
            foreach_expected,
        );

        let while_continue_expected = if n <= 1 { 1 } else { n - 2 };
        assert_int(
            &format!(
                "$sum = 0; $i = 0; do {{ $i += 1; if ($i === 2) {{ continue; }} $sum += $i; }} while ($i < {n}); echo $sum;"
            ),
            while_continue_expected,
        );

        let for_skip_last_expected = (n - 1) * (n - 2) / 2;
        assert_int(
            &format!(
                "$sum = 0; for ($i = 0; $i < {n}; $i++) {{ if ($i === {n} - 1) {{ break; }} $sum += $i; }} echo $sum;"
            ),
            for_skip_last_expected,
        );

        assert_int(
            &format!(
                "$rows = 0; $pairs = [1,2,3,4,5]; foreach ($pairs as $value) {{ if ($value > {n}) {{ break; }} $rows++; }} echo $rows;"
            ),
            if n >= 5 { 5 } else { n } as i64,
        );
    }
}

#[test]
fn control_flow_nested_break_continue_levels_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$count = 0;
for ($i = 0; $i < 3; $i++) {
    for ($j = 0; $j < 3; $j++) {
        if ($i === 1 && $j === 1) {
            continue 2;
        }
        $count++;
    }
}
echo $count;
"#
        ),
        vec!["7".to_string()]
    );
}

#[test]
fn control_flow_nested_break_two_levels_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$count = 0;
for ($i = 0; $i < 3; $i++) {
    for ($j = 0; $j < 3; $j++) {
        $count++;
        if ($i === 1 && $j === 1) {
            break 2;
        }
    }
}
echo $count;
"#
        ),
        vec!["5".to_string()]
    );
}

#[test]
fn control_flow_while_continue_on_odd_numbers_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$i = 0;
$sum = 0;
while ($i < 6) {
    $i++;
    if ($i % 2 === 0) {
        continue;
    }
    $sum += $i;
}
echo $sum;
"#
        ),
        vec!["9".to_string()]
    );
}

#[test]
fn control_flow_switch_fallthrough_like_pattern_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$n = 2;
$out = '';
switch ($n) {
    case 1:
        $out .= 'one';
        break;
    case 2:
        $out .= 'two';
    case 3:
        $out .= '|three';
        break;
    default:
        $out .= '|other';
}
echo $out;
"#
        ),
        vec!["two|three".to_string()]
    );
}

#[test]
fn control_flow_if_elseif_else_chain_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$x = 7;
if ($x < 0) {
    echo 'neg';
} elseif ($x < 5) {
    echo 'small';
} elseif ($x < 10) {
    echo 'mid';
} else {
    echo 'big';
}
"#
        ),
        vec!["mid".to_string()]
    );
}

#[test]
fn control_flow_if_assignment_in_condition_changes_value_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$i = 0;
if (($i = $i + 1) === 1) {
    echo "ok";
}
echo "|";
echo $i;
"#
        ),
        vec!["ok|1".to_string()]
    );
}

#[test]
fn control_flow_switch_uses_loose_comparison_for_matching() {
    assert_eq!(
        run_prints(
            r#"<?php
$input = "0";
$matched = "none";
switch ($input) {
    case 0:
        $matched = "numeric";
        break;
    case "0":
        $matched = "strict_string";
        break;
    default:
        $matched = "other";
}
echo $matched;
"#
        ),
        vec!["numeric".to_string()]
    );
}

#[test]
fn control_flow_while_continue_then_break_breaks_after_sum_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$sum = 0;
$i = 0;
while ($i < 5) {
    $i++;
    if ($i === 2 || $i === 4) {
        continue;
    }
    $sum += $i;
    if ($sum > 4) {
        break;
    }
}
echo $sum;
"#
        ),
        vec!["9".to_string()]
    );
}

#[test]
fn control_flow_for_loop_initial_and_post_update_variants_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$sum = 0;
for ($i = 0, $j = 0; $i < 4; $i++, $j += 2) {
    $sum += $j;
    if ($i === 2) {
        continue;
    }
    $sum += 1;
}
echo $sum;
"#
        ),
        vec!["15".to_string()]
    );
}

#[test]
fn control_flow_conditional_precedence_false_branch_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = false;
$b = true;
if ($a || $b && false) {
    echo "then";
} else {
    echo "else";
}
"#
        ),
        vec!["else".to_string()]
    );
}

#[test]
fn control_flow_nested_if_elseif_chain_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$role = 'viewer';
$label = '';
if ($role === 'admin') {
    $label = 'admin';
} elseif ($role === 'editor') {
    $label = 'editor';
} else {
    $label = 'viewer';
}
if ($label === 'viewer') {
    $label .= ':readonly';
}
echo $label;
"#
        ),
        vec!["viewer:readonly".to_string()]
    );
}

#[test]
fn control_flow_assignment_condition_in_elseif_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$chosen = '';
if (false) {
    $chosen = 'a';
} elseif ($chosen = 'from_elseif') {
    echo $chosen . '|';
}
echo 'done';
"#
        ),
        vec!["from_elseif|done".to_string()]
    );
}

#[test]
fn control_flow_switch_default_in_middle_of_cases_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$mode = 'beta';
$out = '';
switch ($mode) {
    case 'alpha':
        $out .= 'a';
        break;
    default:
        $out .= 'd';
        break;
    case 'beta':
        $out .= 'b';
        break;
}
echo $out;
"#
        ),
        vec!["b".to_string()]
    );
}

#[test]
fn control_flow_do_while_continue_skips_iteration_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$i = 0;
$sum = 0;
do {
    $i++;
    if ($i % 3 === 0) {
        continue;
    }
    $sum += $i;
} while ($i < 6);
echo $sum;
"#
        ),
        vec!["12".to_string()]
    );
}

#[test]
fn control_flow_while_condition_update_in_expression_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$i = 0;
$sum = 0;
while (($i += 1) <= 3) {
    $sum += $i;
}
echo $sum;
"#
        ),
        vec!["6".to_string()]
    );
}

#[test]
fn control_flow_if_short_circuit_skips_rhs_call_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$calls = 0;
$pred = function() use (&$calls) { $calls++; return true; };
$sink = function() use (&$calls) { $calls++; return false; };
if (false && $pred()) {
    echo 'hit';
} else {
    echo $calls;
}
echo '|';
if (true || $sink()) {
    echo 'right';
}
echo '|' . $calls;
"#
        ),
        vec!["0|right|1".to_string()]
    );
}

#[test]
fn control_flow_foreach_break_and_continue_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$values = [1, 2, 3, 4];
$sum = 0;
$i = 0;
foreach ($values as $n) {
    $i++;
    if ($n === 2) {
        continue;
    }
    if ($n === 4) {
        break;
    }
    $sum += $n;
}
echo $sum . '|' . $i;
"#
        ),
        vec!["4|4".to_string()]
    );
}

#[test]
fn control_flow_switch_subject_evaluated_once_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$calls = 0;
switch (($calls++ + 1)) {
    case 1:
        break;
    case 2:
        break;
    default:
        break;
}
echo $calls;
"#
        ),
        vec!["1".to_string()]
    );
}

#[test]
fn control_flow_match_truthy_subject_vs_falsey_default_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo match (false) {
    true => 'T',
    false => 'F',
    default => 'D' } . '|' .
match (0) {
    '' => 'E',
    0 => 'Z',
    default => 'D' };
"#
        ),
        vec!["F|Z".to_string()]
    );
}

#[test]
fn control_flow_computed_switch_subject_is_evaluated_once_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$calls = 0;
$subject = function() use (&$calls) {
    $calls++;
    return 1;
};
switch ($subject()) {
    case 1:
        break;
    default:
        break;
}
echo $calls;
"#
        ),
        vec!["1".to_string()]
    );
}

#[test]
fn control_flow_match_as_elseif_equivalent_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$score = 72;
$out = match (true) {
    $score >= 90 => 'A',
    $score >= 80 => 'B',
    $score >= 70 => 'C',
    default => 'D' };
if ($out === 'C') {
    echo 'pass';
} else {
    echo 'fail';
}
"#
        ),
        vec!["pass".to_string()]
    );
}

#[test]
fn control_flow_elseif_with_assignment_and_reuse_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$x = 0;
if (($x = 10) > 5) {
    echo $x . '|';
} elseif (($x = 20) > 5) {
    echo $x . '|';
}
if ($x === 10) {
    echo 'updated';
}
"#
        ),
        vec!["10|updated".to_string()]
    );
}
