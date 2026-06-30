//! `debug_backtrace` frame shape — function names, depth, and options.

crate::php_cases! {
    backtrace_reports_caller_function_name => {
        r#"<?php
function inner(): string {
    return debug_backtrace()[1]['function'];
}
function outer(): string {
    return inner();
}
echo outer();
"#,
        ["outer"]
    };

    backtrace_current_frame_is_inner_function => {
        r#"<?php
function who(): string {
    return debug_backtrace()[0]['function'];
}
echo who();
"#,
        ["who"]
    };

    backtrace_depth_increases_with_nested_calls => {
        r#"<?php
function c(): int { return count(debug_backtrace()); }
function b(): int { return c(); }
function a(): int { return b(); }
echo a() >= 3 ? 'deep' : 'shallow';
"#,
        ["deep"]
    };

    backtrace_limit_zero_returns_at_least_one_frame => {
        r#"<?php
function one(): int { return count(debug_backtrace(0, 0)); }
echo one() >= 1 ? 'has-frame' : 'empty';
"#,
        ["has-frame"]
    };

    backtrace_limit_one_returns_single_frame => {
        r#"<?php
function single(): int { return count(debug_backtrace(0, 1)); }
echo single() === 1 ? 'one' : 'many';
"#,
        ["one"]
    };

    backtrace_ignore_args_strips_argument_list => {
        r#"<?php
function argCount(): int {
    $frame = debug_backtrace(DEBUG_BACKTRACE_IGNORE_ARGS, 1)[0];
    return count($frame['args'] ?? []);
}
function wrap(int $a, string $b): int {
    return argCount();
}
echo wrap(1, 'x');
"#,
        ["0"]
    };

    backtrace_provides_file_key_in_frame => {
        r#"<?php
function hasFile(): string {
    $f = debug_backtrace()[0]['file'] ?? '';
    return $f !== '' ? 'file' : 'none';
}
echo hasFile();
"#,
        ["file"]
    };

    backtrace_class_name_present_in_method_frame => {
        r#"<?php
class Tracer {
    public function mark(): string {
        return debug_backtrace()[0]['class'] ?? 'none';
    }
}
echo (new Tracer())->mark();
"#,
        ["Tracer"]
    };

    backtrace_type_is_object_for_instance_method => {
        r#"<?php
class Tracer {
    public function mark(): string {
        return debug_backtrace()[0]['type'] ?? '?';
    }
}
echo (new Tracer())->mark();
"#,
        ["->"]
    };

    backtrace_type_is_static_for_static_method => {
        r#"<?php
class Tracer {
    public static function mark(): string {
        return debug_backtrace()[0]['type'] ?? '?';
    }
}
echo Tracer::mark();
"#,
        ["::"]
    };

    backtrace_inside_closure_reports_closure_function => {
        r#"<?php
$fn = function (): string {
    return debug_backtrace()[0]['function'];
};
echo str_contains($fn(), '{closure}') ? 'closure' : $fn();
"#,
        ["closure"]
    };

    backtrace_from_nested_closure_depth => {
        r#"<?php
$outer = function () {
    $inner = function (): int {
        return count(debug_backtrace());
    };
    return $inner();
};
echo $outer() >= 2 ? 'nested' : 'flat';
"#,
        ["nested"]
    };

    backtrace_line_number_is_positive_integer => {
        r#"<?php
function line(): int {
    return debug_backtrace()[0]['line'] ?? 0;
}
echo line() > 0 ? 'line' : 'zero';
"#,
        ["line"]
    };

    backtrace_from_called_helper_identifies_helper => {
        r#"<?php
function helper(): string { return debug_backtrace()[0]['function']; }
function runner(): string { return helper(); }
echo runner();
"#,
        ["helper"]
    };

    backtrace_after_try_catch_still_available => {
        r#"<?php
function safe(): string {
    try { throw new Exception('x'); } catch (Exception $e) { /* handled */ }
    return debug_backtrace()[0]['function'];
}
echo safe();
"#,
        ["safe"]
    };

    backtrace_from_generator_function => {
        r#"<?php
function gen(): Generator {
    yield debug_backtrace()[0]['function'];
}
echo gen()->current();
"#,
        ["gen"]
    };
}
