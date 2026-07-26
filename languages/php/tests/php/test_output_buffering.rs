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
        ["outer-inner-end"]
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
        ["|kept"]
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

    ob_start_with_transform_callback => {
        r#"<?php
ob_start(fn(string $buf): string => strtoupper($buf));
echo 'hello';
echo ob_get_clean();
"#,
        ["HELLO"]
    };

    ob_nested_clean_then_flush_keeps_outer => {
        r#"<?php
ob_start();
echo 'outer-';
ob_start(fn(string $buf): string => '[' . $buf . ']');
echo 'inner';
$inner = ob_get_clean();
ob_get_clean();
echo $inner;
"#,
        ["outer-[inner]"]
    };

    ob_get_flush_applies_callback => {
        r#"<?php
ob_start(fn(string $buf): string => $buf . '-C');
echo 'x';
$flushed = ob_get_flush();
if ($flushed !== false) {
    echo $flushed;
}
"#,
        ["x-Cx"]
    };

    ob_get_flush_without_active_buffer_returns_false => {
        r#"<?php
echo ob_get_flush() === false ? 'false' : 'not';
"#,
        ["false"]
    };

    ob_list_handlers_before_and_after_start => {
        r#"<?php
$before = ob_list_handlers();
ob_start();
$after = ob_list_handlers();
ob_end_clean();
echo (is_array($before) ? count($before) : 0) . '|' . (is_array($after) ? count($after) : 0);
"#,
        ["0|1"]
    };

    ob_end_clean_discarded_and_reusable => {
        r#"<?php
ob_start();
echo 'discard';
ob_end_clean();
ob_start();
echo 'keep';
echo ob_get_clean();
"#,
        ["keep"]
    };

    ob_implicit_flush_controls_runtime_behavior => {
        r#"<?php
ob_start();
echo 'x';
ob_implicit_flush(true);
ob_end_flush();
ob_start();
echo 'y';
ob_end_clean();
echo 'done';
"#,
        ["xydone"]
    };

    ob_status_reports_multiple_buffers => {
        r#"<?php
ob_start();
ob_start();
$status = ob_get_status(true);
echo is_array($status) && array_key_exists(0, $status) && is_array($status[0]) ? 'one' : 'none';
ob_end_clean();
ob_end_clean();
"#,
        ["one"]
    };

    ob_start_with_chunk_size_and_no_flush_flag => {
        r#"<?php
ob_start(null, 1024, false);
echo 'chunk';
echo ob_get_contents();
ob_end_clean();
"#,
        ["chunkchunk"]
    };

    ob_callback_can_strip_tags => {
        r#"<?php
ob_start(fn(string $buf): string => strip_tags($buf));
echo '<b>ok</b> <i>yes</i>';
echo ob_get_clean();
"#,
        ["ok yes"]
    };

    ob_start_with_empty_buffer_reports_zero_length => {
        r#"<?php
ob_start();
echo ob_get_length();
ob_end_clean();
"#,
        ["0"]
    };

    ob_get_contents_after_flush_is_empty => {
        r#"<?php
ob_start();
echo 'hi';
ob_flush();
echo ob_get_contents();
ob_end_clean();
"#,
        ["hi"]
    };

    ob_end_clean_when_no_buffer_is_false => {
        r#"<?php
echo ob_end_clean() ? 'closed' : 'false';
"#,
        ["false"]
    };

    ob_flush_without_contents_preserves_outer_output => {
        r#"<?php
echo 'outer-start-';
ob_start();
echo 'inner';
ob_flush();
echo '-outer-end';
ob_end_clean();
"#,
        ["outer-start-inner-outer-end"]
    };

    ob_start_and_ob_get_clean_preserve_newlines => {
        r#"<?php
ob_start();
echo "a\nb";
echo str_replace("\n", "|", ob_get_clean());
"#,
        ["a|b"]
    };

    ob_nested_levels_with_midlevel_clean => {
        r#"<?php
echo 'L0-';
ob_start();
echo 'L1';
ob_start();
echo 'L2';
$l2 = ob_get_clean();
echo ':' . $l2 . ':';
$l1 = ob_get_clean();
echo $l1;
"#,
        ["L0-:L2:L1"]
    };

    ob_ob_get_clean_discards_outer_when_nested => {
        r#"<?php
ob_start();
echo 'outer';
ob_start();
echo 'inner';
ob_end_clean();
$value = ob_get_clean();
echo $value;
"#,
        ["outer"]
    };

    ob_get_length_after_nested_operations => {
        r#"<?php
ob_start();
echo 'abc';
ob_start();
echo 'de';
echo $inner_len = ob_get_length();
ob_end_flush();
echo '|';
echo $outer_len = ob_get_length();
echo '|';
echo ob_get_clean();
"#,
        ["2|5|abcde"]
    };

    ob_callback_drops_output_when_empty => {
        r#"<?php
ob_start(fn(string $buf): string => '');
echo 'abc';
ob_end_clean();
echo 'after';
"#,
        ["after"]
    };

    ob_implicit_flush_false_retains_buffered_output => {
        r#"<?php
ob_start();
ob_implicit_flush(false);
echo 'x';
echo ob_get_length();
ob_end_clean();
echo 'done';
"#,
        ["1|done"]
    };

    ob_end_flush_with_no_handler => {
        r#"<?php
ob_start(function(string $buf): string { return $buf . 'X'; });
echo 'A';
echo ob_end_flush();
"#,
        ["AX1"]
    };

    ob_status_one_level_flag_false => {
        r#"<?php
ob_start();
$status = ob_get_status();
echo is_array($status) ? 'arr' : 'no';
ob_end_clean();
"#,
        ["arr"]
    };

    ob_get_contents_after_clean_returns_empty => {
        r#"<?php
ob_start();
echo 'tmp';
ob_clean();
echo ob_get_contents();
ob_end_clean();
echo 'end';
"#,
        ["end"]
    };

    ob_list_handlers_length_tracks_start => {
        r#"<?php
ob_start();
ob_start(function($buf) { return strtoupper($buf); });
$handlers = ob_list_handlers();
ob_end_clean();
ob_end_clean();
echo is_array($handlers) ? count($handlers) : 0;
"#,
        ["2"]
    };

    ob_get_flush_with_active_buffer_returns_content => {
        r#"<?php
ob_start();
echo 'xy';
echo ob_get_flush();
"#,
        ["xyxy"]
    };

    ob_start_without_callback_default_gets_buffer => {
        r#"<?php
ob_start();
echo 'abc';
echo ob_get_clean();
"#,
        ["abcabc"]
    };

    ob_start_with_identity_callback_preserves_content => {
        r#"<?php
ob_start(fn(string $buf): string => $buf);
echo 'z';
echo ob_get_clean();
"#,
        ["zz"]
    };

    ob_end_clean_nested_discard_outer_buffer => {
        r#"<?php
ob_start();
echo 'outer';
ob_start();
echo 'inner';
ob_end_clean();
echo '|';
ob_end_clean();
"#,
        ["|"]
    };

    ob_flush_without_content_returns_empty_string => {
        r#"<?php
ob_start();
echo ob_get_length();
echo '|';
ob_flush();
echo ob_get_length();
ob_end_clean();
"#,
        ["0|0"]
    };

    ob_get_contents_after_clean_empty => {
        r#"<?php
ob_start();
echo 'temp';
ob_clean();
echo ob_get_contents() === '' ? 'clean' : 'dirty';
ob_end_clean();
"#,
        ["clean"]
    };

    ob_end_flush_returns_true_on_active_buffer => {
        r#"<?php
ob_start();
echo 'ok';
echo ob_end_flush();
"#,
        ["ok1"]
    };

    ob_start_with_static_handler_array => {
        r#"<?php
class BufferFilters {
    public static function frame(string $buf): string { return '[' . $buf . ']'; }
}
ob_start([BufferFilters::class, 'frame']);
echo 'payload';
$out = ob_get_clean();
echo $out;
"#,
        ["[payload]"]
    };

    ob_start_with_callable_function_name => {
        r#"<?php
ob_start('str_rot13');
echo 'uryyb';
echo ob_get_clean();
"#,
        ["hello"]
    };

    ob_get_length_after_nested_start => {
        r#"<?php
ob_start();
ob_start();
echo 'nested';
$status = ob_get_length();
ob_end_clean();
echo $status;
"#,
        ["6|6"]
    };

    ob_get_status_true_reports_nested_flags => {
        r#"<?php
ob_start();
ob_start();
$status = ob_get_status(true);
echo is_array($status) && count($status) >= 2 ? 'yes' : 'no';
ob_end_clean();
ob_end_clean();
"#,
        ["yes"]
    };

    ob_get_clean_after_end_flush_returns_false => {
        r#"<?php
ob_start();
echo 'flush-me';
ob_end_flush();
echo ob_get_clean() === false ? 'closed' : 'open';
"#,
        ["flush-meclosed"]
    };

    ob_implicit_flush_true_does_not_disable_ob_start => {
        r#"<?php
ob_start();
ob_implicit_flush(true);
echo 'z';
echo ob_get_level() >= 1 ? 'active' : 'off';
ob_end_clean();
"#,
        ["active"]
    };

    ob_start_capture_then_nested_clean => {
        r#"<?php
ob_start();
ob_start();
echo 'inner';
ob_clean();
echo 'outer';
$inner = ob_get_clean();
echo $inner;
"#,
        ["outer"]
    };

    ob_callback_lowercase_output => {
        r#"<?php
ob_start(fn(string $buf): string => strtoupper(strtolower($buf)));
echo 'MiXeD';
echo ob_get_clean();
"#,
        ["MIXED"]
    };

    ob_get_status_depth_false => {
        r#"<?php
ob_start();
ob_start();
echo ob_get_status()['level'];
ob_end_clean();
ob_end_clean();
"#,
        ["2"]
    };

    ob_get_status_without_active_buffer_empty => {
        r#"<?php
$status = ob_get_status();
echo is_array($status) ? 'array' : 'no';
"#,
        ["array"]
    };

    ob_clean_then_get_length_zero => {
        r#"<?php
ob_start();
echo 'temp';
ob_clean();
echo ob_get_length() === 0 ? 'zero' : 'not';
ob_end_clean();
"#,
        ["zero"]
    };

    ob_end_flush_then_get_level_zero => {
        r#"<?php
ob_start();
echo 'x';
ob_end_flush();
echo ob_get_level();
"#,
        ["x0"]
    };

    ob_default_handler_outputs_exactly_once => {
        r#"<?php
ob_start();
echo 'one';
echo 'two';
echo 'three';
echo ob_get_clean();
"#,
        ["onetwothreeonetwothree"]
    };

    ob_nested_buffers_report_level_changes => {
        r#"<?php
ob_start();
echo 'A';
ob_start();
echo 'B';
echo ob_get_level() . '|';
ob_get_clean();
echo ob_get_level();
ob_end_clean();
"#,
        ["2|1"]
    };

    ob_clean_keeps_buffer_handle_open => {
        r#"<?php
ob_start();
echo 'abc';
ob_clean();
echo ob_get_level();
echo '|';
ob_get_clean();
"#,
        ["1|"]
    };

    ob_end_clean_returns_false_after_close => {
        r#"<?php
ob_start();
ob_end_clean();
echo ob_end_clean() ? 'open' : 'closed';
"#,
        ["closed"]
    };

    ob_get_contents_returns_false_without_active_buffer => {
        r#"<?php
echo ob_get_contents() === false ? 'false' : 'true';
"#,
        ["false"]
    };

    ob_callback_add_suffix_and_trim => {
        r#"<?php
ob_start(fn(string $buf): string => trim($buf) . '|ok');
echo '  value  ';
echo ob_get_clean();
"#,
        ["value|ok"]
    };

    ob_get_status_without_level_key_returns_map => {
        r#"<?php
ob_start();
$status = ob_get_status(false);
echo is_array($status) ? (array_key_exists('level', $status) ? 'ok' : 'nolvl') : 'bad';
ob_end_clean();
"#,
        ["ok"]
    };

    ob_get_status_after_end_clean => {
        r#"<?php
ob_start();
echo 'a';
ob_end_clean();
$status = ob_get_status(false);
echo is_array($status) ? 'arr' : 'bad';
echo '|';
echo $status['level'];
"#,
        ["arr|0"]
    };

    ob_get_clean_drops_nested_inner_when_called_in_inner => {
        r#"<?php
echo 'base';
ob_start();
echo 'one';
ob_start();
echo 'two';
echo '|' . ob_get_clean();
echo ob_get_clean();
"#,
        ["base|two|one"]
    };

    ob_clean_when_no_buffer_false => {
        r#"<?php
echo ob_clean() ? 'ok' : 'false';
"#,
        ["false"]
    };

    ob_end_flush_with_return_true_when_active => {
        r#"<?php
ob_start();
echo 'z';
echo ob_end_flush() ? 'ok' : 'bad';
echo ob_get_level();
"#,
        ["zok0"]
    };

    ob_get_length_with_multi_byte => {
        r#"<?php
ob_start();
echo "éclair";
echo ob_get_length();
echo '|';
echo strlen("éclair");
ob_end_clean();
"#,
        ["6|7"]
    };

    ob_start_with_output_handler_array => {
        r#"<?php
ob_start(new class {
    public function __invoke(string $buf): string {
        return str_replace('a', 'A', $buf);
    }
});
echo 'java';
echo ob_get_clean();
        "#,
        ["jAvA"]
    };

    ob_get_clean_without_buffer_returns_false => {
        r#"<?php
echo ob_get_clean() === false ? 'false' : 'not';
"#,
        ["false"]
    };

    ob_get_flush_with_binary_contents => {
        r#"<?php
ob_start();
echo "A";
ob_start();
echo "B";
echo ob_get_flush();
ob_end_clean();
echo ob_get_length();
"#,
        ["B0"]
    };

    ob_end_flush_after_clean_returns_false => {
        r#"<?php
ob_start();
echo 'x';
ob_end_clean();
echo ob_end_flush() ? 'closed' : 'false';
"#,
        ["false"]
    };

    ob_level_after_clean_vs_end_clean => {
        r#"<?php
$base = ob_get_level();
ob_start();
echo ob_get_level() === $base + 1 ? 'nested' : 'wrong';
ob_clean();
echo '|' . (ob_get_level() === $base + 1 ? 'still' : 'gone');
ob_end_clean();
echo '|' . (ob_get_level() === $base ? 'done' : 'open');
"#,
        ["nested|still|done"]
    };

    ob_list_handlers_after_clears => {
        r#"<?php
ob_start();
ob_start(function(string $b): string { return '[' . $b . ']'; });
$before = count(ob_list_handlers());
ob_clean();
$after = count(ob_list_handlers());
ob_end_clean();
ob_end_clean();
echo $before . '|' . $after;
"#,
        ["2|2"]
    };

    ob_get_status_map_with_buffer_closed => {
        r#"<?php
ob_start();
echo 'x';
ob_get_status();
ob_end_clean();
$status = ob_get_status();
echo is_array($status) && array_key_exists('level', $status) ? 'arr' : 'not';
echo '|' . $status['level'];
"#,
        ["arr|0"]
    };

    ob_callback_can_return_empty_string => {
        r#"<?php
ob_start(function(string $buf): string { return ''; });
echo 'drop';
echo ob_get_clean();
"#,
        [""]
    };

    ob_callback_exception_passthrough_not_masked => {
        r#"<?php
ob_start(function(string $buf): string {
    if ($buf === 'boom') { throw new Exception('cb'); }
    return strtoupper($buf);
});
try {
    echo 'boom';
    ob_get_clean();
    echo 'no';
} catch (Exception $e) {
    echo 'caught';
}
"#,
        ["caught"]
    };

    ob_start_with_numeric_chunk_and_non_flush_flag => {
        r#"<?php
ob_start(null, 1, false);
echo 'x';
echo ob_get_length();
echo '|';
echo ob_get_flush() === false ? 'closed' : 'not';
ob_end_clean();
"#,
        ["1|closed"]
    };
}
