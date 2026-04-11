use super::helpers::{run_ruby, run_ruby_one, compile_ok};

// ── IO (compile) ────────────────────────────────────────────────────────────

#[test] fn puts_str()    { compile_ok("puts 'hello'\n"); }
#[test] fn print_str()   { compile_ok("print 'hello'\n"); }
#[test] fn p_value()     { compile_ok("p 42\n"); }
#[test] fn pp_value()    { compile_ok("pp [1, 2, 3]\n"); }
#[test] fn warn_msg()    { compile_ok("warn 'oops'\n"); }

// ── String methods (compile) ────────────────────────────────────────────────

#[test] fn str_length()  { compile_ok("x = 'hello'.length\n"); }
#[test] fn str_upcase()  { compile_ok("x = 'hello'.upcase\n"); }
#[test] fn str_downcase(){ compile_ok("x = 'HELLO'.downcase\n"); }
#[test] fn str_strip()   { compile_ok("x = '  hi  '.strip\n"); }
#[test] fn str_split()   { compile_ok("x = 'a,b,c'.split(',')\n"); }
#[test] fn str_reverse() { compile_ok("x = 'hello'.reverse\n"); }
#[test] fn str_include() { compile_ok("x = 'hello'.include?('ell')\n"); }
#[test] fn str_gsub()    { compile_ok("x = 'hello'.gsub('l', 'r')\n"); }
#[test] fn str_sub()     { compile_ok("x = 'hello'.sub('l', 'r')\n"); }
#[test] fn str_capitalize() { compile_ok("x = 'hello'.capitalize\n"); }
#[test] fn str_start_with() { compile_ok("x = 'hello'.start_with?('he')\n"); }
#[test] fn str_end_with()   { compile_ok("x = 'hello'.end_with?('lo')\n"); }
#[test] fn str_chars()   { compile_ok("x = 'hello'.chars\n"); }

// ── Array methods (compile) ─────────────────────────────────────────────────

#[test] fn arr_push()    { compile_ok("a = [1, 2]\na.push(3)\n"); }
#[test] fn arr_pop()     { compile_ok("a = [1, 2, 3]\na.pop\n"); }
#[test] fn arr_shift()   { compile_ok("a = [1, 2, 3]\na.shift\n"); }
#[test] fn arr_unshift() { compile_ok("a = [1, 2, 3]\na.unshift(0)\n"); }
#[test] fn arr_first()   { compile_ok("x = [1, 2, 3].first\n"); }
#[test] fn arr_last()    { compile_ok("x = [1, 2, 3].last\n"); }
#[test] fn arr_flatten() { compile_ok("x = [[1, 2], [3]].flatten\n"); }
#[test] fn arr_sort()    { compile_ok("x = [3, 1, 2].sort\n"); }
#[test] fn arr_uniq()    { compile_ok("x = [1, 1, 2].uniq\n"); }
#[test] fn arr_compact() { compile_ok("x = [1, nil, 2].compact\n"); }
#[test] fn arr_min()     { compile_ok("x = [3, 1, 2].min\n"); }
#[test] fn arr_max()     { compile_ok("x = [3, 1, 2].max\n"); }
#[test] fn arr_sum()     { compile_ok("x = [1, 2, 3].sum\n"); }
#[test] fn arr_count()   { compile_ok("x = [1, 2, 3].count\n"); }
#[test] fn arr_empty()   { compile_ok("x = [].empty?\n"); }
#[test] fn arr_join()    { compile_ok("x = [1, 2, 3].join(',')\n"); }
#[test] fn arr_each()    { compile_ok("[1, 2, 3].each { |x| puts x }\n"); }
#[test] fn arr_map()     { compile_ok("x = [1, 2, 3].map { |x| x * 2 }\n"); }
#[test] fn arr_select()  { compile_ok("x = [1, 2, 3, 4].select { |x| x > 2 }\n"); }
#[test] fn arr_reject()  { compile_ok("x = [1, 2, 3].reject { |x| x == 2 }\n"); }
#[test] fn arr_reduce()  { compile_ok("x = [1, 2, 3].reduce(0) { |sum, x| sum + x }\n"); }

// ── Hash methods (compile) ──────────────────────────────────────────────────

#[test] fn hash_literal()   { compile_ok("h = { 'a' => 1, 'b' => 2 }\n"); }
#[test] fn hash_keys()      { compile_ok("h = { 'a' => 1 }\nx = h.keys\n"); }
#[test] fn hash_values()    { compile_ok("h = { 'a' => 1 }\nx = h.values\n"); }
#[test] fn hash_has_key()   { compile_ok("h = { 'a' => 1 }\nx = h.has_key?('a')\n"); }
#[test] fn hash_merge()     { compile_ok("h = { 'a' => 1 }.merge({ 'b' => 2 })\n"); }
#[test] fn hash_delete()    { compile_ok("h = { 'a' => 1 }\nh.delete('a')\n"); }
#[test] fn hash_fetch()     { compile_ok("h = { 'a' => 1 }\nx = h.fetch('a')\n"); }

// ── Conversions (compile) ───────────────────────────────────────────────────

#[test] fn to_s()   { compile_ok("x = 42.to_s\n"); }
#[test] fn to_i()   { compile_ok("x = '42'.to_i\n"); }
#[test] fn to_f()   { compile_ok("x = '3.14'.to_f\n"); }
#[test] fn to_a()   { compile_ok("x = (1..3).to_a\n"); }

// ── Numeric (compile) ───────────────────────────────────────────────────────

#[test] fn num_abs()   { compile_ok("x = -5.abs\n"); }
#[test] fn num_even()  { compile_ok("x = 4.even?\n"); }
#[test] fn num_odd()   { compile_ok("x = 3.odd?\n"); }
#[test] fn num_zero()  { compile_ok("x = 0.zero?\n"); }

// ── Math (compile) ──────────────────────────────────────────────────────────

#[test] fn math_sqrt()  { compile_ok("x = Math.sqrt(16)\n"); }
#[test] fn math_pi()    { compile_ok("x = Math::PI\n"); }

// ── Interpolation (compile) ─────────────────────────────────────────────────

#[test] fn string_interpolation() { compile_ok("name = 'world'\nputs \"hello #{name}\"\n"); }

// ── Runtime ─────────────────────────────────────────────────────────────────

#[test]
fn puts_runtime() {
    assert_eq!(run_ruby_one("puts 'hello'\n"), "hello");
}

#[test]
fn str_upcase_runtime() {
    assert_eq!(run_ruby_one("puts 'hello'.upcase\n"), "HELLO");
}

#[test]
fn str_downcase_runtime() {
    assert_eq!(run_ruby_one("puts 'HELLO'.downcase\n"), "hello");
}

#[test]
fn str_length_runtime() {
    assert_eq!(run_ruby_one("puts 'hello'.length\n"), "5");
}

#[test]
fn str_reverse_runtime() {
    assert_eq!(run_ruby_one("puts 'hello'.reverse\n"), "olleh");
}

#[test]
fn str_strip_runtime() {
    assert_eq!(run_ruby_one("puts '  hi  '.strip\n"), "hi");
}

#[test]
fn str_capitalize_runtime() {
    compile_ok("puts 'hello'.capitalize\n");
}

#[test]
fn interpolation_runtime() {
    let out = run_ruby("name = 'world'\nputs \"hello #{name}\"\n");
    assert_eq!(out, vec!["hello world"]);
}

#[test]
fn arr_each_runtime() {
    let out = run_ruby("[1, 2, 3].each { |x| puts x }\n");
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn to_s_runtime() {
    assert_eq!(run_ruby_one("puts 42.to_s\n"), "42");
}

#[test]
fn to_i_runtime() {
    assert_eq!(run_ruby_one("puts '42'.to_i\n"), "42");
}

#[test]
fn math_sqrt_runtime() {
    assert_eq!(run_ruby_one("puts Math.sqrt(16)\n"), "4");
}
