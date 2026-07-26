use super::helpers::compile_ok;

// ── goto statement (PHP 5.3+) ─────────────────────────────────
// Legacy flow control — rarely used but part of the PHP language spec.

#[test]
fn goto_basic() {
    compile_ok(
        r#"<?php
goto end;
echo "skipped";
end:
echo "reached";
"#,
    );
}

#[test]
fn goto_skip_block() {
    compile_ok(
        r#"<?php
$x = 1;
if ($x > 0) { goto positive; }
echo "non-positive";
goto done;
positive:
echo "positive";
done:
"#,
    );
}

#[test]
fn goto_in_loop_exit() {
    compile_ok(
        r#"<?php
$sum = 0;
for ($i = 1; $i <= 10; $i++) {
    if ($i > 5) goto done;
    $sum += $i;
}
done:
echo $sum;
"#,
    );
}

#[test]
fn goto_forward_jump() {
    compile_ok(
        r#"<?php
echo "start";
goto step3;
echo "step2"; // skipped
step3:
echo "step3";
"#,
    );
}

#[test]
fn goto_conditional() {
    compile_ok(
        r#"<?php
$flag = true;
$result = '';
if ($flag) goto found;
$result = 'not found';
goto done;
found:
$result = 'found';
done:
echo $result;
"#,
    );
}

#[test]
fn goto_multiple_labels() {
    compile_ok(
        r#"<?php
$step = 2;
goto {"step$step"};
step1: echo "1"; goto end;
step2: echo "2"; goto end;
step3: echo "3"; goto end;
end:
"#,
    );
}

#[test]
fn goto_nested_function() {
    compile_ok(
        r#"<?php
function process(bool $skip): string {
    if ($skip) goto done;
    $result = 'processed';
    goto end;
    done:
    $result = 'skipped';
    end:
    return $result;
}
echo process(false);
echo process(true);
"#,
    );
}

#[test]
fn goto_error_handling_pattern() {
    compile_ok(
        r#"<?php
function riskyOp(int $n): string {
    if ($n < 0) goto error;
    if ($n === 0) goto zero;
    return "positive: $n";
    error:
    return "error: negative";
    zero:
    return "zero";
}
echo riskyOp(5);
echo riskyOp(0);
echo riskyOp(-1);
"#,
    );
}

#[test]
fn goto_cleanup_pattern() {
    compile_ok(
        r#"<?php
$cleanup_needed = false;
$result = 'pending';
$data = [1, 2, -1, 3];
foreach ($data as $v) {
    if ($v < 0) {
        $cleanup_needed = true;
        goto cleanup;
    }
    $result = "ok:$v";
}
goto done;
cleanup:
$result = "cleaned up";
done:
echo $result;
"#,
    );
}

// ── Runtime `goto` flow (`php_cases!`) ──────────────────────────

crate::php_cases! {
    goto_skips_unreachable_echo => {
        r#"<?php
goto end;
echo 'skip';
end:
echo 'ok';
"#,
        ["ok"]
    };

    goto_positive_branch_label => {
        r#"<?php
$x = 1;
if ($x > 0) goto pos;
echo 'neg';
goto done;
pos:
echo 'pos';
done:
"#,
        ["pos"]
    };

    goto_exits_loop_early => {
        r#"<?php
$sum = 0;
for ($i = 1; $i <= 10; $i++) {
    if ($i > 3) goto stop;
    $sum += $i;
}
stop:
echo $sum;
"#,
        ["6"]
    };

    goto_forward_to_cleanup_block => {
        r#"<?php
$result = 'pending';
foreach ([1, 2, -1] as $v) {
    if ($v < 0) goto cleanup;
    $result = "ok:$v";
    goto done;
}
cleanup:
$result = 'cleaned';
done:
echo $result;
"#,
        ["ok:1"]
    };

    goto_backward_repeat_once => {
        r#"<?php
$i = 0;
again:
echo $i;
$i++;
if ($i < 2) goto again;
"#,
        // echo emits no newline → the two iterations concatenate (verified vs php).
        ["01"]
    };

    goto_in_if_else_branches => {
        r#"<?php
$flag = true;
$result = '';
if ($flag) {
    goto positive;
} else {
    goto negative;
}
positive:
$result .= 'yes';
goto done;
negative:
$result .= 'no';
done:
echo $result;
"#,
        ["yes"]
    };

    goto_from_foreach_with_counter_break => {
        r#"<?php
$sum = 0;
$i = 0;
foreach ([1, 2, 3] as $v) {
    if ($v === 2) {
        $sum += 5;
        goto after_loop;
    }
    $sum += $v;
}
after_loop:
echo $sum;
"#,
        ["6"]
    };

    goto_to_top_like_manual_loop => {
        r#"<?php
$i = 0;
$sum = 0;
loop:
$sum += $i;
$i++;
if ($i <= 3) {
    goto loop;
}
echo $sum;
"#,
        ["6"]
    };

    goto_label_after_try_catch_block => {
        r#"<?php
$result = '';
try {
    throw new Exception('skip');
} catch (Exception $e) {
    $result .= 'caught';
}
goto done;
$result .= ' never';
done:
echo $result;
"#,
        ["caught"]
    };

    goto_nested_condition_branches => {
        r#"<?php
$value = 0;
$out = '';
if (true) {
    goto left;
} else {
    $out .= 'bad';
}
left:
if (false) {
    goto skip;
}
$out .= 'left';
skip:
echo $out;
"#,
        ["left"]
    };

    goto_in_nested_foreach => {
        r#"<?php
$sum = 0;
foreach ([1, 2, 3] as $v) {
    if ($v === 2) {
        $sum += 5;
        goto after_foreach;
    }
    $sum += $v;
}
after_foreach:
echo $sum;
"#,
        ["6"]
    };

    goto_with_switch_control => {
        r#"<?php
$state = 'x';
$out = '';
switch ($state) {
    case 'x':
        $out .= 'x';
        goto done;
    case 'y':
        $out .= 'y';
        break;
}
done:
echo $out;
"#,
        ["x"]
    };

    goto_reentry_loop_without_infinite_spin => {
        r#"<?php
$i = 0;
$out = '';
start:
$out .= $i;
$i++;
if ($i < 3) {
    goto start;
}
echo $out;
"#,
        ["012"]
    };
}
