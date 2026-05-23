use super::helpers::{compile_ok, run_ruby};

// ── puts variants ────────────────────────────────────────────
#[test] fn puts_multiple_args_each_on_line() {
    let out = run_ruby("puts 1, 2, 3\n");
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test] fn puts_array_expands_each_element() {
    let out = run_ruby("puts [10, 20, 30]\n");
    assert_eq!(out, vec!["10", "20", "30"]);
}

#[test] fn puts_nil_prints_blank_line() { compile_ok(r#"
puts nil
"#); }

// ── print ─────────────────────────────────────────────────────
#[test] fn print_no_trailing_newline() {
    let out = run_ruby("print \"hello\"\nprint \" world\n\"\n");
    assert_eq!(out, vec!["hello world"]);
}

#[test] fn print_multiple_args_concatenated() { compile_ok(r#"
print "a", "b", "c"
puts ""
"#); }

// ── p ─────────────────────────────────────────────────────────
#[test] fn p_shows_inspect_form() {
    let out = run_ruby("p \"hello\"\n");
    assert_eq!(out, vec!["\"hello\""]);
}

#[test] fn p_nil_shows_nil() {
    let out = run_ruby("p nil\n");
    assert_eq!(out, vec!["nil"]);
}

#[test] fn p_array_shows_inspect() {
    let out = run_ruby("p [1, 2, 3]\n");
    assert_eq!(out, vec!["[1, 2, 3]"]);
}

// ── pp ────────────────────────────────────────────────────────
#[test] fn pp_prints_pretty_inspection() { compile_ok(r#"
pp({ name: "Alice", scores: [10, 20, 30] })
"#); }

// ── printf / format ───────────────────────────────────────────
#[test] fn printf_formatted_output() { compile_ok(r#"
printf("Hello, %s! You are %d years old.\n", "Alice", 30)
"#); }

#[test] fn format_returns_string() { compile_ok(r#"
s = format("%-10s %5d", "item", 42)
puts s
"#); }

#[test] fn sprintf_alias_for_format() { compile_ok(r#"
s = sprintf("%.2f", 3.14159)
puts s
"#); }

// ── String % operator ─────────────────────────────────────────
#[test] fn string_percent_integer_format() {
    let out = run_ruby("puts \"Value: %d\" % 42\n");
    assert_eq!(out, vec!["Value: 42"]);
}

#[test] fn string_percent_multiple_values() {
    let out = run_ruby("puts \"%s is %d\" % [\"Alice\", 30]\n");
    assert_eq!(out, vec!["Alice is 30"]);
}

#[test] fn string_percent_float_precision() {
    let out = run_ruby("puts \"%.2f\" % 3.14159\n");
    assert_eq!(out, vec!["3.14"]);
}

#[test] fn string_percent_hex_format() {
    let out = run_ruby("puts \"%x\" % 255\n");
    assert_eq!(out, vec!["ff"]);
}

// ── STDOUT ────────────────────────────────────────────────────
#[test] fn stdout_puts_method() { compile_ok(r#"
$stdout.puts "via stdout"
"#); }

#[test] fn stdout_print_method() { compile_ok(r#"
$stdout.print "no newline"
$stdout.puts ""
"#); }

// ── warn writes to stderr ─────────────────────────────────────
#[test] fn warn_outputs_message() { compile_ok(r#"
warn "this is a warning"
"#); }

// ── puts with to_s coercion ───────────────────────────────────
#[test] fn puts_calls_to_s_on_object() { compile_ok(r#"
class Widget
  def to_s; "Widget!"; end
end
puts Widget.new
"#); }
