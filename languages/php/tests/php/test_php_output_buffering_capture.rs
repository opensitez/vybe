use super::helpers::run_prints;

// Every caller below passes a whole multi-statement PROGRAM, not a single
// expression, so the source runs as-is. Prefixing `echo ` (as these helpers
// used to) turned the leading `ob_start();` into `echo ob_start();`, which
// echoes `true` INTO the buffer it just opened — an invisible extra "1" in
// every expectation, visible only when the buffer was flushed rather than
// discarded.
fn assert_int(src: &str, expected: i64) {
    assert_eq!(
        run_prints(&format!("<?php {} ", src)),
        vec![expected.to_string()]
    );
}

fn assert_str(src: &str, expected: &str) {
    assert_eq!(
        run_prints(&format!("<?php {} ", src)),
        vec![expected.to_string()]
    );
}

#[test]
fn php_output_buffering_capture() {
    for n in 1..=20_i64 {
        assert_int(
            &format!(
                "ob_start();\n$payload = str_repeat('x', {n});\nob_start();\nob_start();\necho $payload;\n$inner = ob_get_clean();\nob_end_clean();\necho strlen($inner);\nob_end_flush();"
            ),
            n,
        );

        let mut levels = String::new();
        for _ in 0..n {
            levels.push_str("ob_start();\n");
        }
        let mut close = String::new();
        for _ in 0..n {
            close.push_str("ob_end_clean();\n");
        }
        assert_int(
            &format!(
                "ob_start();\n{levels}$payload = str_repeat('y', 1);\n$level = ob_get_level();\n{close}echo $level;\nob_end_flush();",
                levels = levels,
                close = close,
            ),
            n + 1,
        );
    }
}

#[test]
fn php_output_buffering_ob_get_contents_and_clean() {
    assert_str(
        "ob_start(); echo 'hello'; $inner = ob_get_contents(); ob_clean(); echo $inner; ob_end_flush();",
        "hello",
    );

    assert_str(
        "ob_start(); echo 'start'; $before = ob_get_level(); $c = ob_get_clean(); $after = ob_get_level(); echo $before . ':' . $c . ':' . $after;",
        "1:start:0",
    );
}

#[test]
fn php_output_buffering_nested_levels_and_flush_flow() {
    assert_str(
        "ob_start(); echo 'a'; ob_start(); echo 'b'; $l1 = ob_get_level(); $inner = ob_get_clean(); echo $l1 . ':' . $inner . ':' . ob_get_contents(); ob_end_flush();",
        "a2:b:a",
    );

    assert_str(
        "ob_start(); echo 'outer'; ob_start(); echo 'inner'; ob_end_flush(); echo ob_get_length();",
        "outerinner10",
    );
}

#[test]
fn php_output_buffering_clean_vs_end_clean_semantics() {
    assert_str(
        "ob_start(); echo 'kept'; ob_clean(); echo 'x'; $level = ob_get_level(); ob_end_clean(); echo $level;",
        "1",
    );

    assert_str(
        "ob_start(); echo 'first'; $v1 = ob_get_length(); $v2 = (int)ob_end_clean(); $v3 = ob_get_level(); echo $v1 . ':' . $v2 . ':' . $v3;",
        "5:1:0",
    );
}

#[test]
fn php_output_buffering_get_level_after_nested_failures() {
    assert_str(
        "ob_start(); $before = ob_get_level(); ob_start(); echo 'x'; ob_end_flush(); $after = ob_get_level(); ob_end_clean(); echo $before . ':' . $after;",
        "1:1",
    );

    assert_str(
        "echo 'base'; $level = ob_get_level(); echo ':' . $level;",
        "base:0",
    );
}

#[test]
fn php_output_buffering_callback_handler() {
    assert_str(
        "ob_start(function(string $chunk) { return strtoupper($chunk); }); echo 'ab'; $v = ob_get_clean(); echo $v;",
        "AB",
    );

    assert_str(
        "ob_start(fn(string $chunk): string => strrev($chunk)); echo 'abc'; echo ob_get_clean();",
        "cba",
    );
}

#[test]
fn php_output_buffering_end_flush_returns_bool() {
    assert_str(
        "ob_start(); echo 'ping'; $ok = ob_end_flush(); echo ':' . ($ok ? '1' : '0');",
        "ping:1",
    );
}

#[test]
fn php_output_buffering_list_handlers_order() {
    assert_str(
        "ob_start(); echo 'A'; ob_start(); echo 'B'; $handlers = ob_list_handlers(); $size = count($handlers); ob_end_flush(); ob_end_flush(); echo $size;",
        "AB2",
    );
}
