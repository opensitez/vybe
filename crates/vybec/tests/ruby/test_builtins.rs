use vybec::parser_ruby::parse;
use vybec::compiler_ruby::Compiler;

fn compile_ok(src: &str) {
    let program = parse(src).expect("parse failed");
    let mut c = Compiler::new();
    let res = c.compile(&program);
    assert!(res.is_ok(), "compile failed for:\n{}\nerror: {:?}", src, res.err());
}

// ── IO ─────────────────────────────────────────────────────
#[test] fn puts_basic() { compile_ok("puts 'hello'"); }
#[test] fn puts_multi() { compile_ok("puts 'hello', 'world'"); }
#[test] fn print_basic() { compile_ok("print 'hello'"); }
#[test] fn p_inspect() { compile_ok("p 42"); }
#[test] fn gets_input() { compile_ok("x = gets"); }

// ── String methods ─────────────────────────────────────────
#[test] fn str_length() { compile_ok("x = 'hello'.length"); }
#[test] fn str_upcase() { compile_ok("x = 'hello'.upcase"); }
#[test] fn str_downcase() { compile_ok("x = 'HELLO'.downcase"); }
#[test] fn str_strip() { compile_ok("x = '  hello  '.strip"); }
#[test] fn str_lstrip() { compile_ok("x = '  hello'.lstrip"); }
#[test] fn str_rstrip() { compile_ok("x = 'hello  '.rstrip"); }
#[test] fn str_include() { compile_ok("x = 'hello'.include?('ell')"); }
#[test] fn str_start_with() { compile_ok("x = 'hello'.start_with?('hel')"); }
#[test] fn str_end_with() { compile_ok("x = 'hello'.end_with?('llo')"); }
#[test] fn str_index() { compile_ok("x = 'hello'.index('l')"); }
#[test] fn str_gsub() { compile_ok("x = 'hello'.gsub('l', 'r')"); }
#[test] fn str_sub() { compile_ok("x = 'hello'.sub('l', 'r')"); }
#[test] fn str_split() { compile_ok("x = 'a,b,c'.split(',')"); }
#[test] fn str_split_default() { compile_ok("x = 'hello world'.split"); }
#[test] fn str_chars() { compile_ok("x = 'hello'.chars"); }
#[test] fn str_reverse() { compile_ok("x = 'hello'.reverse"); }

// ── Conversions ────────────────────────────────────────────
#[test] fn to_s() { compile_ok("x = 42.to_s"); }
#[test] fn to_i() { compile_ok("x = '42'.to_i"); }
#[test] fn to_f() { compile_ok("x = '3.14'.to_f"); }
#[test] fn integer_conv() { compile_ok("x = Integer('42')"); }
#[test] fn float_conv() { compile_ok("x = Float('3.14')"); }
#[test] fn string_conv() { compile_ok("x = String(42)"); }

// ── Array methods ──────────────────────────────────────────
#[test] fn arr_push() { compile_ok("a = [1, 2]\na.push(3)"); }
#[test] fn arr_pop() { compile_ok("a = [1, 2, 3]\na.pop"); }
#[test] fn arr_shift() { compile_ok("a = [1, 2, 3]\na.shift"); }
#[test] fn arr_first() { compile_ok("a = [1, 2, 3]\nx = a.first"); }
#[test] fn arr_last() { compile_ok("a = [1, 2, 3]\nx = a.last"); }
#[test] fn arr_length() { compile_ok("a = [1, 2, 3]\nx = a.length"); }
#[test] fn arr_count() { compile_ok("a = [1, 2, 3]\nx = a.count"); }
#[test] fn arr_empty() { compile_ok("a = []\nx = a.empty?"); }
#[test] fn arr_reverse() { compile_ok("a = [1, 2, 3].reverse"); }
#[test] fn arr_sort() { compile_ok("a = [3, 1, 2].sort"); }
#[test] fn arr_min() { compile_ok("a = [3, 1, 2].min"); }
#[test] fn arr_max() { compile_ok("a = [3, 1, 2].max"); }
#[test] fn arr_sum() { compile_ok("a = [1, 2, 3].sum"); }
#[test] fn arr_join() { compile_ok("a = ['a', 'b', 'c'].join(', ')"); }
#[test] fn arr_join_default() { compile_ok("a = ['a', 'b', 'c'].join"); }
#[test] fn arr_index() { compile_ok("a = [1, 2, 3]\nx = a[1]"); }

// ── Hash methods ───────────────────────────────────────────
#[test] fn hash_keys() { compile_ok("h = {a: 1, b: 2}\nk = h.keys"); }
#[test] fn hash_values() { compile_ok("h = {a: 1, b: 2}\nv = h.values"); }
#[test] fn hash_has_key() { compile_ok("h = {a: 1}\nx = h.has_key?(:a)"); }
#[test] fn hash_merge() { compile_ok("h = {a: 1}.merge({b: 2})"); }
#[test] fn hash_fetch() { compile_ok("h = {a: 1}\nx = h.fetch(:a)"); }

// ── Enumerable (block-taking) ─────────────────────────────
#[test] fn each_block() { compile_ok("[1, 2, 3].each { |x| puts x }"); }
#[test] fn map_block() { compile_ok("doubled = [1, 2, 3].map { |x| x * 2 }"); }
#[test] fn select_block() { compile_ok("evens = [1, 2, 3, 4].select { |x| x % 2 == 0 }"); }
#[test] fn reject_block() { compile_ok("odds = [1, 2, 3, 4].select { |x| x % 2 != 0 }"); }
#[test] fn reduce_block() { compile_ok("sum = [1, 2, 3].reduce { |acc, x| acc + x }"); }
#[test] fn reduce_initial() { compile_ok("sum = [1, 2, 3].reduce(0) { |acc, x| acc + x }"); }
#[test] fn any_block() { compile_ok("x = [1, 2, 3].any? { |n| n > 2 }"); }
#[test] fn all_block() { compile_ok("x = [1, 2, 3].all? { |n| n > 0 }"); }
#[test] fn flat_map_block() { compile_ok("x = [[1, 2], [3, 4]].flat_map { |a| a }"); }

// ── Nil check ──────────────────────────────────────────────
#[test] fn nil_check() { compile_ok("x = nil\nputs x.nil?"); }

// ── Math builtins ──────────────────────────────────────────
#[test] fn math_sqrt() { compile_ok("x = sqrt(16)"); }
#[test] fn math_abs() { compile_ok("x = abs(-5)"); }
#[test] fn math_rand() { compile_ok("x = rand"); }

// ── Sleep ──────────────────────────────────────────────────
#[test] fn sleep_call() { compile_ok("sleep(1)"); }
