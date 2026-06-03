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
