//! `Console.Write` / `Console.WriteLine` over proper WASI I/O
//! (`wasi:cli/stdout` + `wasi:io/streams.blocking-write-and-flush`), NOT the
//! line-oriented `wasi:logging`.
//!
//! The distinction under test: `Write` emits its text with NO trailing
//! newline, so consecutive `Write`s land on one line; `WriteLine` appends
//! exactly one `\n`. The harness concatenates every stdout fragment and then
//! splits on `\n`, so a captured "line" is real output between newlines —
//! which is exactly what makes the Write-vs-WriteLine difference observable.

use super::helpers::run_csharp;

#[test]
fn writeline_appends_newline_so_each_call_is_its_own_line() {
    assert_eq!(
        run_csharp(
            r#"// console_write
Console.WriteLine("a"); Console.WriteLine("b");"#
        ),
        &["a", "b"]
    );
}

#[test]
fn write_emits_no_newline_so_consecutive_writes_share_a_line() {
    assert_eq!(
        run_csharp(
            r#"// console_write
Console.Write("a"); Console.Write("b"); Console.WriteLine();"#
        ),
        &["ab"]
    );
}

#[test]
fn write_then_writeline_join_on_the_same_line() {
    assert_eq!(
        run_csharp(
            r#"// console_write
Console.Write("a"); Console.WriteLine("b");"#
        ),
        &["ab"]
    );
}

#[test]
fn lone_write_without_trailing_newline_is_still_captured() {
    assert_eq!(
        run_csharp(
            r#"// console_write
Console.Write("solo");"#
        ),
        &["solo"]
    );
}

#[test]
fn writeline_no_args_emits_a_blank_line() {
    assert_eq!(
        run_csharp(
            r#"// console_write
Console.WriteLine();"#
        ),
        &[""]
    );
}

#[test]
fn write_then_empty_writeline_terminates_the_started_line() {
    assert_eq!(
        run_csharp(
            r#"// console_write
Console.Write("x"); Console.WriteLine();"#
        ),
        &["x"]
    );
}

#[test]
fn writeline_bool_capitalises_true() {
    assert_eq!(
        run_csharp(
            r#"// console_write
Console.WriteLine(true);"#
        ),
        &["True"]
    );
}

#[test]
fn write_bool_capitalises_false_without_newline() {
    assert_eq!(
        run_csharp(
            r#"// console_write
Console.Write(false); Console.WriteLine("!");"#
        ),
        &["False!"]
    );
}

#[test]
fn writeline_null_string_prints_empty_line() {
    assert_eq!(
        run_csharp(
            r#"// console_write
Console.WriteLine((string)null);"#
        ),
        &[""]
    );
}

#[test]
fn writeline_integer_renders_digits() {
    assert_eq!(
        run_csharp(
            r#"// console_write
Console.WriteLine(42);"#
        ),
        &["42"]
    );
}

#[test]
fn write_integer_then_write_integer_concatenate() {
    assert_eq!(
        run_csharp(
            r#"// console_write
Console.Write(1); Console.Write(2); Console.Write(3); Console.WriteLine();"#
        ),
        &["123"]
    );
}

#[test]
fn interleaved_write_and_writeline_across_multiple_lines() {
    assert_eq!(
        run_csharp(
            r#"Console.Write("a"); Console.Write("b"); Console.WriteLine("c"); Console.WriteLine("d");"#
        ),
        &["abc", "d"]
    );
}

#[test]
fn writeline_empty_string_is_a_blank_line_between_content() {
    assert_eq!(
        run_csharp(
            r#"// console_write
Console.WriteLine("a"); Console.WriteLine(""); Console.WriteLine("b");"#
        ),
        &["a", "", "b"]
    );
}

#[test]
fn write_in_a_loop_builds_one_line() {
    assert_eq!(
        run_csharp(r#"for (int i = 0; i < 3; i++) { Console.Write(i); } Console.WriteLine();"#),
        &["012"]
    );
}
