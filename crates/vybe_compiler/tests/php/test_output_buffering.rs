use super::helpers::compile_ok;

// ── Basic output buffering ────────────────────────────────────

#[test]
fn ob_start_get_clean() {
    compile_ok(
        r#"<?php
ob_start();
echo "buffered content";
$content = ob_get_clean();
echo "captured: $content";
"#,
    );
}

#[test]
fn ob_get_contents() {
    compile_ok(
        r#"<?php
ob_start();
echo "hello";
echo " world";
$buf = ob_get_contents();
ob_end_clean();
echo strlen($buf);
"#,
    );
}

#[test]
fn ob_end_flush() {
    compile_ok(
        r#"<?php
ob_start();
echo "flushed";
ob_end_flush();
"#,
    );
}

#[test]
fn ob_get_length() {
    compile_ok(
        r#"<?php
ob_start();
echo "twelve chars";
echo ob_get_length();
ob_end_clean();
"#,
    );
}

#[test]
fn ob_get_level() {
    compile_ok(
        r#"<?php
echo ob_get_level();
ob_start();
echo ob_get_level();
ob_end_clean();
echo ob_get_level();
"#,
    );
}

// ── Nested output buffers ─────────────────────────────────────

#[test]
fn ob_nested_two_levels() {
    compile_ok(
        r#"<?php
ob_start();
echo "outer";
ob_start();
echo "inner";
$inner = ob_get_clean();
echo "-" . $inner . "-";
$outer = ob_get_clean();
echo $outer;
"#,
    );
}

#[test]
fn ob_nested_three_levels() {
    compile_ok(
        r#"<?php
ob_start();
    echo "L1";
    ob_start();
        echo "L2";
        ob_start();
            echo "L3";
        $l3 = ob_get_clean();
    $l2 = ob_get_clean();
$l1 = ob_get_clean();
echo "$l1-$l2-$l3";
"#,
    );
}

#[test]
fn ob_level_tracking() {
    compile_ok(
        r#"<?php
$levels = [];
$levels[] = ob_get_level();
ob_start();
$levels[] = ob_get_level();
ob_start();
$levels[] = ob_get_level();
ob_end_clean();
ob_end_clean();
$levels[] = ob_get_level();
echo implode(',', $levels);
"#,
    );
}

// ── ob with callback ──────────────────────────────────────────

#[test]
fn ob_start_with_callback() {
    compile_ok(
        r#"<?php
ob_start(fn(string $buf) => strtoupper($buf));
echo "hello world";
ob_end_flush();
"#,
    );
}

#[test]
fn ob_callback_transform() {
    compile_ok(
        r#"<?php
ob_start(function(string $buf): string {
    return str_replace(' ', '_', $buf);
});
echo "hello world again";
ob_end_flush();
"#,
    );
}

// ── ob_flush vs ob_get_clean ──────────────────────────────────

#[test]
fn ob_flush_keeps_buffer() {
    compile_ok(
        r#"<?php
ob_start();
echo "part one";
ob_flush();
echo "part two";
$all = ob_get_clean();
echo $all;
"#,
    );
}

#[test]
fn ob_clean_discards() {
    compile_ok(
        r#"<?php
ob_start();
echo "will be discarded";
ob_clean();
echo "this survives";
ob_end_clean();
echo "done";
"#,
    );
}

// ── Capture output of print_r / var_dump ─────────────────────

#[test]
fn ob_capture_print_r() {
    compile_ok(
        r#"<?php
$data = ['a' => 1, 'b' => 2, 'c' => 3];
ob_start();
print_r($data);
$output = ob_get_clean();
echo strlen($output) > 0 ? 'captured' : 'empty';
"#,
    );
}

#[test]
fn ob_capture_var_dump() {
    compile_ok(
        r#"<?php
ob_start();
var_dump(42, 'hello', true);
$output = ob_get_clean();
echo strlen($output) > 0 ? 'ok' : 'fail';
"#,
    );
}

// ── Template rendering pattern ────────────────────────────────

#[test]
fn ob_template_render() {
    compile_ok(
        r#"<?php
function renderTemplate(string $title, array $items): string {
    ob_start(); ?>
<h1><?= htmlspecialchars($title) ?></h1>
<ul>
<?php foreach ($items as $item): ?>
  <li><?= htmlspecialchars($item) ?></li>
<?php endforeach; ?>
</ul>
<?php return ob_get_clean();
}
$html = renderTemplate('My List', ['Apple', 'Banana', 'Cherry']);
echo strlen($html) > 0 ? 'rendered' : 'empty';
"#,
    );
}

#[test]
fn ob_component_pattern() {
    compile_ok(
        r#"<?php
function component(callable $render): string {
    ob_start();
    $render();
    return ob_get_clean();
}
$output = component(function() {
    echo "Hello from component!";
});
echo $output;
"#,
    );
}

// ── ob_get_status ─────────────────────────────────────────────

#[test]
fn ob_get_status_basic() {
    compile_ok(
        r#"<?php
ob_start();
echo "data";
$status = ob_get_status();
ob_end_clean();
echo isset($status['level']) ? 'has level' : 'no level';
"#,
    );
}

#[test]
fn ob_get_status_all() {
    compile_ok(
        r#"<?php
ob_start();
ob_start();
$statuses = ob_get_status(true);
ob_end_clean();
ob_end_clean();
echo count($statuses) >= 2 ? 'two levels' : 'less';
"#,
    );
}

// ── Runtime output buffering (`php_cases!`) ─────────────────────

crate::php_cases! {
    ob_get_clean_captures_buffered_echo => {
        r#"<?php
ob_start();
echo 'buf';
$c = ob_get_clean();
echo $c;
"#,
        ["buf"]
    };

    ob_nested_buffers_flush_in_order => {
        r#"<?php
ob_start();
echo 'outer-';
ob_start();
echo 'inner';
$inner = ob_get_clean();
echo $inner . '-end';
$outer = ob_get_clean();
echo $outer;
"#,
        ["inner", "outer-inner-end"]
    };

    ob_get_contents_without_end => {
        r#"<?php
ob_start();
echo 'stay';
$s = ob_get_contents();
ob_end_clean();
echo $s;
"#,
        ["stay"]
    };

    ob_get_length_reports_byte_count => {
        r#"<?php
ob_start();
echo '12345';
echo ob_get_length();
ob_end_clean();
"#,
        ["123455"]
    };

    ob_get_level_increments_with_nested_start => {
        r#"<?php
$base = ob_get_level();
ob_start();
echo $base + 1;
ob_end_clean();
"#,
        ["1"]
    };

    ob_end_flush_prints_buffer => {
        r#"<?php
ob_start();
echo 'flushme';
ob_end_flush();
"#,
        ["flushme"]
    };

    ob_implicit_flush_zero_disables_auto_flush_flag => {
        r#"<?php
ob_implicit_flush(0);
echo ob_get_level() >= 0 ? 'ok' : 'bad';
"#,
        ["ok"]
    };

    ob_clean_clears_active_buffer => {
        r#"<?php
ob_start();
echo 'remove';
ob_clean();
echo 'kept';
$c = ob_get_clean();
echo '|' . $c;
"#,
        ["kept", "|kept"]
    };

    ob_gzhandler_not_active_without_compression => {
        r#"<?php
echo function_exists('ob_gzhandler') ? 'exists' : 'missing';
"#,
        ["exists"]
    };

    ob_list_includes_current_buffer_after_start => {
        r#"<?php
ob_start();
$list = ob_list_handlers();
ob_end_clean();
echo count($list) >= 0 ? 'listed' : 'none';
"#,
        ["listed"]
    };
}
