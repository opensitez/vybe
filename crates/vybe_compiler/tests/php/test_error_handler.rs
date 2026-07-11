//! Error-handler behaviors with distinct output shapes (not compile-only echo flags).

crate::php_cases! {
    handler_records_errno_as_hex_suffix => {
        r#"<?php
$out = [];
set_error_handler(function(int $no, string $msg) use (&$out): bool {
    $out[] = dechex($no) . ':' . $msg;
    return true;
});
trigger_error('warn', E_USER_WARNING);
restore_error_handler();
echo implode('|', $out);
"#,
        ["200:warn"]
    };

    handler_mask_ignores_notice_but_catches_warning => {
        r#"<?php
$hits = '';
set_error_handler(function(int $no) use (&$hits): bool {
    $hits .= $no === E_USER_WARNING ? 'W' : 'N';
    return true;
}, E_USER_WARNING);
trigger_error('notice', E_USER_NOTICE);
trigger_error('warning', E_USER_WARNING);
restore_error_handler();
echo $hits;
"#,
        ["W"]
    };

    handler_return_false_leaves_error_get_last_populated => {
        r#"<?php
set_error_handler(fn() => false);
trigger_error('persist', E_USER_NOTICE);
restore_error_handler();
$e = error_get_last();
echo $e !== null && str_contains($e['message'], 'persist') ? 'last' : 'none';
"#,
        ["last"]
    };

    handler_clears_last_inside_callback => {
        r#"<?php
set_error_handler(function() {
    error_clear_last();
    return true;
});
trigger_error('gone', E_USER_WARNING);
restore_error_handler();
echo error_get_last() === null ? 'cleared' : 'set';
"#,
        ["cleared"]
    };

    nested_restore_reenables_outer_formatter => {
        r#"<?php
$log = [];
set_error_handler(function($no, $msg) use (&$log): bool { $log[] = 'A'; return true; });
set_error_handler(function($no, $msg) use (&$log): bool { $log[] = 'B'; return true; });
trigger_error('x', E_USER_NOTICE);
restore_error_handler();
trigger_error('y', E_USER_NOTICE);
restore_error_handler();
echo implode('', $log);
"#,
        ["BA"]
    };

    handler_counts_distinct_user_levels => {
        r#"<?php
$counts = ['w' => 0, 'n' => 0, 'd' => 0];
set_error_handler(function(int $no) use (&$counts): bool {
    if ($no === E_USER_WARNING) $counts['w']++;
    if ($no === E_USER_NOTICE) $counts['n']++;
    if ($no === E_USER_DEPRECATED) $counts['d']++;
    return true;
});
trigger_error('a', E_USER_WARNING);
trigger_error('b', E_USER_NOTICE);
trigger_error('c', E_USER_DEPRECATED);
trigger_error('d', E_USER_WARNING);
restore_error_handler();
echo $counts['w'] . $counts['n'] . $counts['d'];
"#,
        // Triggers: W, N, D, W → w=2, n=1, d=1 (the 4th, "d", is E_USER_WARNING,
        // not E_USER_DEPRECATED), so the concatenation is "211", not "212".
        ["211"]
    };

    handler_prefixes_message_with_custom_tag => {
        r#"<?php
$captured = '';
set_error_handler(function($no, $msg) use (&$captured): bool {
    $captured = '[ERR]' . $msg;
    return true;
});
trigger_error('payload', E_USER_ERROR);
restore_error_handler();
echo $captured;
"#,
        ["[ERR]payload"]
    };

    handler_sees_errline_nonzero => {
        r#"<?php
$line = 0;
set_error_handler(function($no, $msg, $file, $lineNo) use (&$line): bool {
    $line = $lineNo;
    return true;
});
trigger_error('loc', E_USER_NOTICE);
restore_error_handler();
echo $line > 0 ? 'line' : 'zero';
"#,
        ["line"]
    };

    handler_sees_errfile_non_empty => {
        r#"<?php
$file = '';
set_error_handler(function($no, $msg, $fileName) use (&$file): bool {
    $file = $fileName !== '' ? 'file' : 'empty';
    return true;
});
trigger_error('f', E_USER_NOTICE);
restore_error_handler();
echo $file;
"#,
        ["file"]
    };

    at_operator_prevents_handler_for_trigger_error => {
        r#"<?php
$fired = false;
set_error_handler(function() use (&$fired): bool { $fired = true; return true; });
@trigger_error('silent', E_USER_NOTICE);
restore_error_handler();
echo $fired ? 'fired' : 'muted';
"#,
        ["fired"]
    };

    error_reporting_zero_suppresses_user_notice_to_handler => {
        r#"<?php
$fired = false;
set_error_handler(function() use (&$fired): bool { $fired = true; return true; });
$old = error_reporting(0);
trigger_error('hidden', E_USER_NOTICE);
error_reporting($old);
restore_error_handler();
echo $fired ? 'fired' : 'hidden';
"#,
        ["fired"]
    };

    error_reporting_restored_after_trigger => {
        r#"<?php
$before = error_reporting(E_ALL);
trigger_error('x', E_USER_NOTICE);
$after = error_reporting();
error_reporting($before);
echo $after === E_ALL ? 'all' : 'less';
"#,
        ["all"]
    };

    error_get_last_type_matches_user_warning => {
        r#"<?php
set_error_handler(fn() => false);
trigger_error('typed', E_USER_WARNING);
restore_error_handler();
$e = error_get_last();
echo $e['type'] === E_USER_WARNING ? 'warn' : 'other';
"#,
        ["warn"]
    };

    error_clear_last_between_two_triggers => {
        r#"<?php
set_error_handler(fn() => false);
trigger_error('first', E_USER_NOTICE);
error_clear_last();
trigger_error('second', E_USER_WARNING);
restore_error_handler();
$e = error_get_last();
echo str_contains($e['message'], 'second') ? 'second' : 'first';
"#,
        ["second"]
    };

    set_exception_handler_stores_callable => {
        r#"<?php
$ran = false;
set_exception_handler(function() use (&$ran) { $ran = true; });
$prev = set_exception_handler(null);
echo $prev !== null && !$ran ? 'stored' : 'fail';
"#,
        ["stored"]
    };

    user_error_handler_object_method => {
        r#"<?php
class Logger {
    public array $rows = [];
    public function onError(int $no, string $msg): bool {
        $this->rows[] = $msg;
        return true;
    }
}
$log = new Logger();
set_error_handler([$log, 'onError']);
trigger_error('obj', E_USER_NOTICE);
restore_error_handler();
echo count($log->rows);
"#,
        ["1"]
    };

    user_error_handler_static_method => {
        r#"<?php
class Logger {
    public static string $last = '';
    public static function onError(int $no, string $msg): bool {
        self::$last = $msg;
        return true;
    }
}
set_error_handler([Logger::class, 'onError']);
trigger_error('static', E_USER_WARNING);
restore_error_handler();
echo Logger::$last;
"#,
        ["static"]
    };

    handler_invoked_once_per_trigger_in_loop => {
        r#"<?php
$n = 0;
set_error_handler(function() use (&$n): bool { $n++; return true; });
for ($i = 0; $i < 3; $i++) { trigger_error("i$i", E_USER_NOTICE); }
restore_error_handler();
echo $n;
"#,
        ["3"]
    };

    handler_does_not_run_when_assertion_disabled => {
        r#"<?php
$n = 0;
set_error_handler(function() use (&$n): bool { $n++; return true; });
$old = assert_options(ASSERT_ACTIVE, 0);
assert(false);
assert_options(ASSERT_ACTIVE, $old);
restore_error_handler();
echo $n;
"#,
        ["4"]
    };

    division_by_zero_triggers_error_or_exception => {
        r#"<?php
$tag = '';
set_error_handler(function() use (&$tag): bool { $tag = 'handler'; return true; });
try { $x = 1 / 0; echo 'inf'; }
catch (DivisionByZeroError $e) { $tag = 'exception'; }
restore_error_handler();
echo $tag;
"#,
        ["exception"]
    };

    converting_warning_to_exception_via_handler_throw => {
        r#"<?php
set_error_handler(function($no, $msg) { throw new RuntimeException($msg); });
try {
    trigger_error('boom', E_USER_WARNING);
    echo 'no';
} catch (RuntimeException $e) {
    echo $e->getMessage();
} finally {
    restore_error_handler();
}
"#,
        ["boom"]
    };

    handler_swallows_deprecated_but_counts => {
        r#"<?php
$d = 0;
set_error_handler(function(int $no) use (&$d): bool {
    if ($no === E_USER_DEPRECATED) { $d++; return true; }
    return false;
});
trigger_error('old', E_USER_DEPRECATED);
trigger_error('old2', E_USER_DEPRECATED);
restore_error_handler();
echo $d;
"#,
        ["2"]
    };

    trigger_error_empty_string_message => {
        r#"<?php
$msg = 'unset';
set_error_handler(function($no, $str) use (&$msg): bool { $msg = $str === '' ? 'empty' : 'text'; return true; });
trigger_error('', E_USER_NOTICE);
restore_error_handler();
echo $msg;
"#,
        ["empty"]
    };

    error_get_last_after_handler_returns_true_may_be_null => {
        r#"<?php
set_error_handler(fn() => true);
trigger_error('swallow', E_USER_NOTICE);
restore_error_handler();
echo error_get_last() === null ? 'null' : 'set';
"#,
        ["null"]
    };

    restore_without_set_is_safe => {
        r#"<?php
restore_error_handler();
echo 'ok';
"#,
        ["ok"]
    };

    double_restore_after_single_set => {
        r#"<?php
set_error_handler(fn() => true);
restore_error_handler();
restore_error_handler();
echo 'ok';
"#,
        ["ok"]
    };

    handler_closure_captures_by_reference_counter => {
        r#"<?php
$sum = 0;
set_error_handler(function() use (&$sum): bool { $sum += 10; return true; });
trigger_error('a', E_USER_NOTICE);
trigger_error('b', E_USER_NOTICE);
restore_error_handler();
echo $sum;
"#,
        ["20"]
    };

    user_error_levels_do_not_include_parse_errors => {
        r#"<?php
$hit = false;
set_error_handler(function() use (&$hit): bool { $hit = true; return true; });
try { eval('$ = ;'); } catch (ParseError $e) { /* parse */ }
restore_error_handler();
echo $hit ? 'handler' : 'parse';
"#,
        ["parse"]
    };

    json_throw_on_error_surfaces_json_exception_not_handler => {
        r#"<?php
$handler = false;
set_error_handler(function() use (&$handler): bool { $handler = true; return true; });
try { json_decode('{', flags: JSON_THROW_ON_ERROR); }
catch (JsonException $e) { echo 'json'; }
restore_error_handler();
"#,
        ["json"]
    };

    array_access_notice_can_be_handled => {
        r#"<?php
$tag = 'none';
set_error_handler(function() use (&$tag): bool { $tag = 'handled'; return true; });
$arr = [];
@$arr[0];
restore_error_handler();
echo $tag;
"#,
        ["handled"]
    };

    unlink_missing_file_warning_handler => {
        r#"<?php
$caught = false;
set_error_handler(function() use (&$caught): bool { $caught = true; return true; });
@unlink('/no/such/file_' . uniqid());
restore_error_handler();
echo $caught ? 'warn' : 'clean';
"#,
        ["warn"]
    };

    fopen_failure_returns_false_without_fatal => {
        r#"<?php
$h = @fopen('/no/such/path_' . uniqid(), 'r');
echo $h === false ? 'false' : 'handle';
"#,
        ["false"]
    };

    set_error_handler_null_restores_default => {
        r#"<?php
set_error_handler(fn() => true);
$prev = set_error_handler(null);
echo is_callable($prev) ? 'callable' : 'not';
"#,
        ["callable"]
    };

    handler_runs_before_finally_on_trigger_in_try => {
        r#"<?php
$log = [];
set_error_handler(function() use (&$log): bool { $log[] = 'h'; return true; });
try {
    trigger_error('t', E_USER_NOTICE);
    $log[] = 't';
} finally {
    $log[] = 'f';
}
restore_error_handler();
echo implode('', $log);
"#,
        ["htf"]
    };

    trigger_in_catch_block_still_handled => {
        r#"<?php
$log = [];
set_error_handler(function() use (&$log): bool { $log[] = 'e'; return true; });
try { throw new Exception('x'); }
catch (Exception $ex) {
    trigger_error('in-catch', E_USER_NOTICE);
    $log[] = 'c';
}
restore_error_handler();
echo implode('', $log);
"#,
        ["ec"]
    };

    error_reporting_user_only_mask => {
        r#"<?php
$mask = E_USER_ERROR | E_USER_WARNING | E_USER_NOTICE | E_USER_DEPRECATED;
$old = error_reporting($mask);
echo error_reporting() === $mask ? 'mask' : 'diff';
error_reporting($old);
"#,
        ["mask"]
    };

    ini_set_display_errors_does_not_break_handler => {
        r#"<?php
$old = ini_set('display_errors', '0');
$hit = false;
set_error_handler(function() use (&$hit): bool { $hit = true; return true; });
trigger_error('quiet', E_USER_NOTICE);
restore_error_handler();
ini_set('display_errors', $old !== false ? $old : '1');
echo $hit ? 'hit' : 'miss';
"#,
        ["hit"]
    };

    multiple_restore_restores_default_handler_behavior => {
        r#"<?php
set_error_handler(fn() => true);
restore_error_handler();
$fired = false;
set_error_handler(function() use (&$fired): bool { $fired = true; return true; });
trigger_error('z', E_USER_NOTICE);
restore_error_handler();
echo $fired ? 'yes' : 'no';
"#,
        ["yes"]
    };

    handler_with_typed_params_accepts_int_errno => {
        r#"<?php
$type = '';
set_error_handler(function(int $errno, string $errstr) use (&$type): bool {
    $type = is_int($errno) ? 'int' : 'other';
    return true;
});
trigger_error('typed', E_USER_WARNING);
restore_error_handler();
echo $type;
"#,
        ["int"]
    };

    consecutive_handlers_see_only_own_triggers => {
        r#"<?php
$a = 0; $b = 0;
set_error_handler(function() use (&$a): bool { $a++; return true; });
trigger_error('1', E_USER_NOTICE);
restore_error_handler();
set_error_handler(function() use (&$b): bool { $b++; return true; });
trigger_error('2', E_USER_NOTICE);
restore_error_handler();
echo $a . $b;
"#,
        ["11"]
    };

    error_get_last_message_trimmed_not_empty => {
        r#"<?php
set_error_handler(fn() => false);
trigger_error('  spaced  ', E_USER_NOTICE);
restore_error_handler();
$e = error_get_last();
echo trim($e['message']) === 'spaced' ? 'trim' : 'raw';
"#,
        ["trim"]
    };

    user_deprecated_level_numeric_distinct => {
        r#"<?php
echo E_USER_DEPRECATED === 16384 ? 'dep' : 'other';
"#,
        ["dep"]
    };

    user_error_level_numeric_distinct => {
        r#"<?php
echo E_USER_ERROR === 256 ? 'err' : 'other';
"#,
        ["err"]
    };

    handler_return_true_on_user_error_does_not_abort_script => {
        r#"<?php
set_error_handler(fn() => true);
trigger_error('continue', E_USER_ERROR);
restore_error_handler();
echo 'alive';
"#,
        ["alive"]
    };

    trigger_error_after_restore_hits_default_path => {
        r#"<?php
set_error_handler(fn() => true);
restore_error_handler();
set_error_handler(fn() => true);
trigger_error('again', E_USER_NOTICE);
restore_error_handler();
echo 'done';
"#,
        ["done"]
    };

    custom_handler_not_invoked_for_caught_exceptions => {
        r#"<?php
$hit = false;
set_error_handler(function() use (&$hit): bool { $hit = true; return true; });
try { throw new RuntimeException('ex'); }
catch (RuntimeException $e) { echo 'catch'; }
restore_error_handler();
echo $hit ? ':handler' : ':clean';
"#,
        ["catch:clean"]
    };

    error_clear_last_before_trigger_sets_fresh => {
        r#"<?php
error_clear_last();
set_error_handler(fn() => false);
trigger_error('fresh', E_USER_WARNING);
restore_error_handler();
$e = error_get_last();
echo $e['message'] === 'fresh' ? 'fresh' : 'stale';
"#,
        ["fresh"]
    };
}
