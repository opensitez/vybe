use super::helpers::{run_ruby, run_ruby_one, compile_ok};

// ── chomp ────────────────────────────────────────────────────────────────────

#[test]
fn str_chomp_removes_newline() {
    compile_ok("x = \"hello\\n\".chomp\n");
}

#[test]
fn str_chomp_runtime() {
    assert_eq!(run_ruby_one("puts \"hello\\n\".chomp\n"), "hello");
}

// ── chop ─────────────────────────────────────────────────────────────────────

#[test]
fn str_chop_removes_last_char() {
    compile_ok("x = 'hello'.chop\n");
}

#[test]
fn str_chop_runtime() {
    assert_eq!(run_ruby_one("puts 'hello'.chop\n"), "hell");
}

// ── lstrip ───────────────────────────────────────────────────────────────────

#[test]
fn str_lstrip_removes_leading_whitespace() {
    compile_ok("x = '   hi'.lstrip\n");
}

#[test]
fn str_lstrip_runtime() {
    assert_eq!(run_ruby_one("puts '   hi'.lstrip\n"), "hi");
}

// ── rstrip ───────────────────────────────────────────────────────────────────

#[test]
fn str_rstrip_removes_trailing_whitespace() {
    compile_ok("x = 'hi   '.rstrip\n");
}

#[test]
fn str_rstrip_runtime() {
    assert_eq!(run_ruby_one("puts 'hi   '.rstrip\n"), "hi");
}

// ── center ───────────────────────────────────────────────────────────────────

#[test]
fn str_center_width() {
    compile_ok("x = 'hi'.center(10)\n");
}

#[test]
fn str_center_with_fill_char() {
    compile_ok("x = 'hi'.center(10, '-')\n");
}

// ── ljust / rjust ─────────────────────────────────────────────────────────────

#[test]
fn str_ljust_left_justify() {
    compile_ok("x = 'hi'.ljust(10)\n");
}

#[test]
fn str_rjust_right_justify() {
    compile_ok("x = 'hi'.rjust(10)\n");
}

#[test]
fn str_rjust_runtime() {
    assert_eq!(run_ruby_one("puts 'hi'.rjust(5)\n"), "   hi");
}

// ── tr ───────────────────────────────────────────────────────────────────────

#[test]
fn str_tr_translate_chars() {
    compile_ok("x = 'hello'.tr('el', 'ip')\n");
}

#[test]
fn str_tr_runtime() {
    assert_eq!(run_ruby_one("puts 'hello'.tr('aeiou', '*')\n"), "h*ll*");
}

// ── tr_s ─────────────────────────────────────────────────────────────────────

#[test]
fn str_tr_s_translate_and_squeeze() {
    compile_ok("x = 'hello'.tr_s('l', 'r')\n");
}

// ── squeeze ──────────────────────────────────────────────────────────────────

#[test]
fn str_squeeze_consecutive_chars() {
    compile_ok("x = 'aaabbbccc'.squeeze\n");
}

#[test]
fn str_squeeze_runtime() {
    assert_eq!(run_ruby_one("puts 'aaabbbccc'.squeeze\n"), "abc");
}

#[test]
fn str_squeeze_specific_chars() {
    compile_ok("x = 'aaabbbccc'.squeeze('a')\n");
}

// ── delete ───────────────────────────────────────────────────────────────────

#[test]
fn str_delete_chars() {
    compile_ok("x = 'hello'.delete('l')\n");
}

#[test]
fn str_delete_runtime() {
    assert_eq!(run_ruby_one("puts 'hello world'.delete('lo')\n"), "he wrd");
}

// ── count ────────────────────────────────────────────────────────────────────

#[test]
fn str_count_matching_chars() {
    compile_ok("x = 'hello world'.count('lo')\n");
}

// ── hex / oct ─────────────────────────────────────────────────────────────────

#[test]
fn str_hex_to_integer() {
    compile_ok("x = 'ff'.hex\n");
}

#[test]
fn str_hex_runtime() {
    assert_eq!(run_ruby_one("puts 'ff'.hex\n"), "255");
}

#[test]
fn str_oct_to_integer() {
    compile_ok("x = '077'.oct\n");
}

#[test]
fn str_oct_runtime() {
    assert_eq!(run_ruby_one("puts '077'.oct\n"), "63");
}

// ── succ / next ───────────────────────────────────────────────────────────────

#[test]
fn str_succ_next_in_sequence() {
    compile_ok("x = 'a'.succ\n");
}

#[test]
fn str_next_alias() {
    compile_ok("x = 'a'.next\n");
}

// ── index / rindex ────────────────────────────────────────────────────────────

#[test]
fn str_index_first_occurrence() {
    compile_ok("x = 'hello'.index('l')\n");
}

#[test]
fn str_index_runtime() {
    assert_eq!(run_ruby_one("puts 'hello'.index('l')\n"), "2");
}

#[test]
fn str_rindex_last_occurrence() {
    compile_ok("x = 'hello'.rindex('l')\n");
}

#[test]
fn str_rindex_runtime() {
    assert_eq!(run_ruby_one("puts 'hello'.rindex('l')\n"), "3");
}

// ── slice ─────────────────────────────────────────────────────────────────────

#[test]
fn str_slice_extract_substring() {
    compile_ok("x = 'hello'.slice(1, 3)\n");
}

#[test]
fn str_slice_runtime() {
    assert_eq!(run_ruby_one("puts 'hello'.slice(1, 3)\n"), "ell");
}

// ── insert ───────────────────────────────────────────────────────────────────

#[test]
fn str_insert_at_position() {
    compile_ok("x = 'hello'.insert(2, 'XY')\n");
}

// ── scan ─────────────────────────────────────────────────────────────────────

#[test]
fn str_scan_find_all_matches() {
    compile_ok("x = 'one two three'.scan('e')\n");
}

// ── match / match? ───────────────────────────────────────────────────────────

#[test]
fn str_match_regex() {
    compile_ok("x = 'hello123'.match(/\\d+/)\n");
}

#[test]
fn str_match_predicate() {
    compile_ok("x = 'hello'.match?(/ell/)\n");
}

// ── =~ operator ───────────────────────────────────────────────────────────────

#[test]
fn str_regex_match_operator() {
    compile_ok("x = 'hello' =~ /ell/\n");
}

// ── encode ───────────────────────────────────────────────────────────────────

#[test]
fn str_encode_change_encoding() {
    compile_ok("x = 'hello'.encode('UTF-8')\n");
}

// ── bytes ─────────────────────────────────────────────────────────────────────

#[test]
fn str_bytes_array() {
    compile_ok("x = 'hi'.bytes\n");
}

// ── bytesize ──────────────────────────────────────────────────────────────────

#[test]
fn str_bytesize_byte_length() {
    compile_ok("x = 'hello'.bytesize\n");
}

#[test]
fn str_bytesize_runtime() {
    assert_eq!(run_ruby_one("puts 'hello'.bytesize\n"), "5");
}

// ── each_char ─────────────────────────────────────────────────────────────────

#[test]
fn str_each_char_with_block() {
    compile_ok("'abc'.each_char { |c| puts c }\n");
}

// ── each_line ─────────────────────────────────────────────────────────────────

#[test]
fn str_each_line_with_block() {
    compile_ok("\"line1\\nline2\\n\".each_line { |l| puts l.chomp }\n");
}

// ── lines ─────────────────────────────────────────────────────────────────────

#[test]
fn str_lines_split_into_array() {
    compile_ok("x = \"one\\ntwo\\nthree\\n\".lines\n");
}

// ── chars ─────────────────────────────────────────────────────────────────────

#[test]
fn str_chars_returning_array() {
    compile_ok("x = 'abc'.chars\n");
}

#[test]
fn str_chars_runtime_count() {
    assert_eq!(run_ruby_one("puts 'hello'.chars.length\n"), "5");
}

// ── % operator ────────────────────────────────────────────────────────────────

#[test]
fn str_percent_sprintf_format() {
    compile_ok("x = 'hello %s' % 'world'\n");
}

// ── format / sprintf ──────────────────────────────────────────────────────────

#[test]
fn str_sprintf_format() {
    compile_ok("x = sprintf('value: %d', 42)\n");
}

#[test]
fn str_format_function() {
    compile_ok("x = format('pi is %.2f', 3.14159)\n");
}

// ── String * repetition ───────────────────────────────────────────────────────

#[test]
fn str_repetition_star() {
    compile_ok("x = 'ha' * 3\n");
}

#[test]
fn str_repetition_runtime() {
    assert_eq!(run_ruby_one("puts 'ha' * 3\n"), "hahaha");
}

// ── freeze / frozen? ──────────────────────────────────────────────────────────

#[test]
fn str_freeze_and_frozen_predicate() {
    compile_ok("x = 'hello'.freeze\ny = x.frozen?\n");
}

// ── dup ───────────────────────────────────────────────────────────────────────

#[test]
fn str_dup_on_frozen() {
    compile_ok("x = 'hello'.freeze\ny = x.dup\n");
}

// ── gsub with regex ───────────────────────────────────────────────────────────

#[test]
fn str_gsub_with_regex_pattern() {
    compile_ok("x = 'hello world'.gsub(/[aeiou]/, '*')\n");
}

// ── sub with regex and capture group ─────────────────────────────────────────

#[test]
fn str_sub_with_capture_group() {
    compile_ok("x = 'hello world'.sub(/(\\w+)/, 'HI')\n");
}

// ── split with regex ─────────────────────────────────────────────────────────

#[test]
fn str_split_with_regex() {
    compile_ok("x = 'one1two2three'.split(/\\d/)\n");
}

// ── split with limit ─────────────────────────────────────────────────────────

#[test]
fn str_split_with_limit() {
    compile_ok("x = 'a,b,c,d'.split(',', 2)\n");
}

#[test]
fn str_split_limit_runtime() {
    let out = run_ruby("x = 'a,b,c,d'.split(',', 2)\nputs x.length\n");
    assert_eq!(out, vec!["2"]);
}

// ── start_with? multiple args ────────────────────────────────────────────────

#[test]
fn str_start_with_multiple_args() {
    compile_ok("x = 'hello'.start_with?('he', 'wo')\n");
}

// ── end_with? multiple args ───────────────────────────────────────────────────

#[test]
fn str_end_with_multiple_args() {
    compile_ok("x = 'hello'.end_with?('lo', 'xx')\n");
}

// ── strip vs lstrip vs rstrip differences ────────────────────────────────────

#[test]
fn str_strip_vs_lstrip_vs_rstrip() {
    compile_ok(r#"
s = "  hello  "
a = s.strip
b = s.lstrip
c = s.rstrip
"#);
}

// ── Heredoc ───────────────────────────────────────────────────────────────────

#[test]
fn str_squiggly_heredoc() {
    compile_ok(r#"
x = <<~HEREDOC
  hello
  world
HEREDOC
"#);
}

// ── %w[] word array literal ───────────────────────────────────────────────────

#[test]
fn str_percent_w_word_array() {
    compile_ok("x = %w[foo bar baz]\n");
}

#[test]
fn str_percent_w_runtime_length() {
    assert_eq!(run_ruby_one("puts %w[foo bar baz].length\n"), "3");
}

// ── %i[] symbol array literal ────────────────────────────────────────────────

#[test]
fn str_percent_i_symbol_array() {
    compile_ok("x = %i[foo bar baz]\n");
}

// ── String <=> comparison ────────────────────────────────────────────────────

#[test]
fn str_spaceship_comparison() {
    compile_ok("x = 'apple' <=> 'banana'\n");
}

#[test]
fn str_spaceship_runtime() {
    assert_eq!(run_ruby_one("puts('apple' <=> 'banana')\n"), "-1");
}

// ── String << append ─────────────────────────────────────────────────────────

#[test]
fn str_shovel_append() {
    compile_ok("s = 'hello'\ns << ' world'\n");
}

#[test]
fn str_shovel_runtime() {
    assert_eq!(run_ruby_one("s = 'hello'\ns << ' world'\nputs s\n"), "hello world");
}
