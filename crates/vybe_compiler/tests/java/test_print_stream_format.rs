//! `java.io.PrintStream` write/format semantics (JLS + java.util.Formatter).
//!
//! Expectations are the real Java stdout observed LINE BY LINE: `print`/
//! `append`/`printf`/`format` write without a newline (buffered into the
//! current line), `println` completes the line. Each captured entry is one
//! complete output line — exactly what `javac && java` prints.

use crate::helpers::run_main;

#[test]
fn print_stream_format_width_precision_integer() {
    let out = run_main(r#"System.out.format("%8.3f", 1.5); System.out.println("done");"#);
    assert_eq!(out, vec!["   1.500done"]);
}

#[test]
fn print_stream_format_left_align_string() {
    let out = run_main(r#"System.out.format("%-6s|", "vybe"); System.out.println("|");"#);
    assert_eq!(out, vec!["vybe  ||"]);
}

#[test]
fn print_stream_format_multiple_specifiers() {
    let out = run_main(r#"System.out.format("%s=%d", "id", 9); System.out.println("");"#);
    assert_eq!(out, vec!["id=9"]);
}

#[test]
fn print_stream_printf_hex_uppercase() {
    let out = run_main(r#"System.out.printf("%X", 255); System.out.println("");"#);
    assert_eq!(out, vec!["FF"]);
}

#[test]
fn print_stream_printf_zero_padded_integer() {
    let out = run_main(r#"System.out.printf("%05d", 7); System.out.println("");"#);
    assert_eq!(out, vec!["00007"]);
}

#[test]
fn print_stream_printf_plus_sign_on_positive() {
    let out = run_main(r#"System.out.printf("%+d", 12); System.out.println("");"#);
    assert_eq!(out, vec!["+12"]);
}

#[test]
fn print_stream_printf_scientific_notation() {
    let out = run_main(r#"System.out.printf("%e", 1000.0); System.out.println("");"#);
    assert_eq!(out, vec!["1.000000e+03"]);
}

#[test]
fn print_stream_printf_grouped_integer() {
    let out = run_main(r#"System.out.printf("%,d", 1234567); System.out.println("");"#);
    assert_eq!(out, vec!["1,234,567"]);
}

#[test]
fn print_stream_append_char_adds_single_character() {
    let out = run_main(r#"System.out.append('A'); System.out.println("");"#);
    assert_eq!(out, vec!["A"]);
}

#[test]
fn print_stream_append_char_sequence() {
    let out = run_main(r#"System.out.append("hel"); System.out.println("lo");"#);
    assert_eq!(out, vec!["hello"]);
}

#[test]
fn print_stream_append_char_sequence_start_end() {
    let out = run_main(r#"System.out.append("hello", 1, 4); System.out.println("");"#);
    assert_eq!(out, vec!["ell"]);
}

#[test]
fn print_stream_append_returns_print_stream_for_chaining() {
    let out =
        run_main(r#"System.out.append("a").append("b").append("c"); System.out.println("");"#);
    assert_eq!(out, vec!["abc"]);
}

#[test]
fn print_stream_append_then_format() {
    let out =
        run_main(r#"System.out.append("[").format("%d", 5).append("]"); System.out.println("");"#);
    assert_eq!(out, vec!["[5]"]);
}

#[test]
fn print_stream_format_then_append() {
    let out = run_main(r#"System.out.format("%d", 3).append("!"); System.out.println("");"#);
    assert_eq!(out, vec!["3!"]);
}

#[test]
fn print_stream_append_int_via_string_conversion() {
    let out = run_main(r#"System.out.append(String.valueOf(42)); System.out.println("");"#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn print_stream_append_null_char_sequence() {
    // JLS PrintStream.append(null) appends the four characters "null".
    let out = run_main(r#"System.out.append((String) null); System.out.println("null");"#);
    assert_eq!(out, vec!["nullnull"]);
}

#[test]
fn print_stream_format_percent_literal() {
    let out = run_main(r#"System.out.format("100%%"); System.out.println("");"#);
    assert_eq!(out, vec!["100%"]);
}

#[test]
fn print_stream_format_newline_escape() {
    // %n terminates the line: stdout is "a\nbc\n" — two lines.
    let out = run_main(r#"System.out.format("a%nb"); System.out.println("c");"#);
    assert_eq!(out, vec!["a", "bc"]);
}

#[test]
fn print_stream_printf_boolean_format() {
    let out = run_main(r#"System.out.printf("%b", true); System.out.println("");"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn print_stream_printf_character_format() {
    let out = run_main(r#"System.out.printf("%c", 90); System.out.println("");"#);
    assert_eq!(out, vec!["Z"]);
}

#[test]
fn print_stream_format_octal_integer() {
    let out = run_main(r#"System.out.format("%o", 8); System.out.println("");"#);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn print_stream_format_hash_hex_alternate_form() {
    let out = run_main(r#"System.out.format("%#x", 10); System.out.println("");"#);
    assert_eq!(out, vec!["0xa"]);
}

#[test]
fn print_stream_append_empty_char_sequence() {
    let out = run_main(r#"System.out.append(""); System.out.println("x");"#);
    assert_eq!(out, vec!["x"]);
}

#[test]
fn print_stream_append_subsequence_at_start() {
    let out = run_main(r#"System.out.append("abcdef", 0, 3); System.out.println("");"#);
    assert_eq!(out, vec!["abc"]);
}

#[test]
fn print_stream_append_subsequence_at_end() {
    let out = run_main(r#"System.out.append("abcdef", 4, 6); System.out.println("");"#);
    assert_eq!(out, vec!["ef"]);
}

#[test]
fn print_stream_format_negative_width_left_align_integer() {
    let out = run_main(r#"System.out.format("%-5d|", 9); System.out.println("");"#);
    assert_eq!(out, vec!["9    |"]);
}

#[test]
fn print_stream_printf_mixed_string_and_integer() {
    let out = run_main(r#"System.out.printf("n=%d", 4); System.out.println("");"#);
    assert_eq!(out, vec!["n=4"]);
}

#[test]
fn print_stream_format_general_format_specifier() {
    // Java %g: 6 significant digits — String.format("%g", 3.5) == "3.50000".
    let out = run_main(r#"System.out.format("%g", 3.5); System.out.println("");"#);
    assert_eq!(out, vec!["3.50000"]);
}

#[test]
fn print_stream_append_char_then_println() {
    let out = run_main(r#"System.out.append('x'); System.out.println('y');"#);
    assert_eq!(out, vec!["xy"]);
}

#[test]
fn print_stream_format_three_arguments_in_order() {
    let out = run_main(r#"System.out.format("%s:%d:%s", "a", 1, "b"); System.out.println("");"#);
    assert_eq!(out, vec!["a:1:b"]);
}

#[test]
fn print_stream_printf_width_on_string() {
    let out = run_main(r#"System.out.printf("%10s", "hi"); System.out.println("");"#);
    assert_eq!(out, vec!["        hi"]);
}

#[test]
fn print_stream_format_precision_truncates_float() {
    let out = run_main(r#"System.out.format("%.1f", 2.56); System.out.println("");"#);
    assert_eq!(out, vec!["2.6"]);
}

#[test]
fn print_stream_append_before_println_without_extra_newline() {
    let out = run_main(r#"System.out.append("pre-"); System.out.println("fix");"#);
    assert_eq!(out, vec!["pre-fix"]);
}

#[test]
fn print_stream_format_returns_print_stream_reference() {
    // JLS: format returns this. The "ok" it writes joins the same line as
    // the println that follows.
    let out = run_main(
        r#"java.io.PrintStream ps = System.out; java.io.PrintStream same = ps.format("%s", "ok"); System.out.println(same == ps);"#,
    );
    assert_eq!(out, vec!["oktrue"]);
}

#[test]
fn print_stream_append_returns_print_stream_reference() {
    let out = run_main(
        r#"java.io.PrintStream ps = System.out; java.io.PrintStream same = ps.append("x"); System.out.println(same == ps);"#,
    );
    assert_eq!(out, vec!["xtrue"]);
}

#[test]
fn print_stream_printf_uppercase_scientific() {
    let out = run_main(r#"System.out.printf("%E", 1000.0); System.out.println("");"#);
    assert_eq!(out, vec!["1.000000E+03"]);
}

#[test]
fn print_stream_format_negative_integer() {
    let out = run_main(r#"System.out.format("%d", -88); System.out.println("");"#);
    assert_eq!(out, vec!["-88"]);
}

#[test]
fn print_stream_append_unicode_char_sequence() {
    let out = run_main(r#"System.out.append("AB"); System.out.println("");"#);
    assert_eq!(out, vec!["AB"]);
}

#[test]
fn print_stream_format_combined_width_and_precision() {
    let out = run_main(r#"System.out.format("%10.2f", 3.1); System.out.println("");"#);
    assert_eq!(out, vec!["      3.10"]);
}

#[test]
fn print_stream_append_char_sequence_with_format_suffix() {
    let out = run_main(r#"System.out.append("val=").format("%d", 2); System.out.println("");"#);
    assert_eq!(out, vec!["val=2"]);
}

#[test]
fn print_stream_printf_indexed_arguments() {
    let out = run_main(r#"System.out.printf("%2$s %1$d", 7, "items"); System.out.println("");"#);
    assert_eq!(out, vec!["items 7"]);
}

#[test]
fn print_stream_format_null_string_as_null_text() {
    let out = run_main(r#"System.out.format("%s", (String) null); System.out.println("");"#);
    assert_eq!(out, vec!["null"]);
}

#[test]
fn print_stream_append_multiple_chars_without_newlines() {
    let out =
        run_main(r#"System.out.append('1').append('2').append('3'); System.out.println("");"#);
    assert_eq!(out, vec!["123"]);
}

#[test]
fn print_stream_format_lowercase_hex() {
    let out = run_main(r#"System.out.format("%x", 16); System.out.println("");"#);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn print_stream_printf_and_append_interleaved() {
    // All three writes land on one line, closed by the println.
    let out = run_main(
        r#"System.out.printf("%d", 1); System.out.append("+"); System.out.printf("%d", 2); System.out.println("");"#,
    );
    assert_eq!(out, vec!["1+2"]);
}

#[test]
fn print_stream_format_empty_format_string() {
    let out = run_main(r#"System.out.format(""); System.out.println("empty");"#);
    assert_eq!(out, vec!["empty"]);
}

#[test]
fn print_stream_append_char_sequence_full_string() {
    let out = run_main(r#"System.out.append("vybe", 0, 4); System.out.println("");"#);
    assert_eq!(out, vec!["vybe"]);
}

#[test]
fn print_stream_format_long_integer() {
    let out = run_main(r#"System.out.format("%d", 1000000L); System.out.println("");"#);
    assert_eq!(out, vec!["1000000"]);
}

#[test]
fn print_stream_append_then_printf_on_same_line() {
    let out = run_main(
        r#"System.out.append("["); System.out.printf("%s", "x"); System.out.println("]");"#,
    );
    assert_eq!(out, vec!["[x]"]);
}

#[test]
fn print_stream_format_space_padding_flag() {
    let out = run_main(r#"System.out.format("% d", 5); System.out.println("");"#);
    assert_eq!(out, vec![" 5"]);
}
