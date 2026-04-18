use vybec::parser_ruby::parse;
use vybec::compiler_ruby::Compiler;

fn compile_ok(src: &str) {
    let program = parse(src).expect("parse failed");
    let c = Compiler::new();
    let res = c.compile(&program);
    assert!(res.is_ok(), "compile failed for:\n{}\nerror: {:?}", src, res.err());
}

fn parse_ok(src: &str) -> bool {
    parse(src).is_ok()
}

// ═══════════════════════════════════════════════════════════
// 1. String#* (repeat)
// ═══════════════════════════════════════════════════════════
#[test] fn str_repeat() { compile_ok("x = 'ha' * 3"); }
#[test] fn str_repeat_var() { compile_ok("s = 'abc'\nx = s * 2"); }

// ═══════════════════════════════════════════════════════════
// 2. Negative indexing
// ═══════════════════════════════════════════════════════════
#[test] fn neg_index_literal() { compile_ok("a = [1, 2, 3]\nx = a[-1]"); }
#[test] fn neg_index_expr() { compile_ok("a = [1, 2, 3]\ni = -2\nx = a[i]"); }

// ═══════════════════════════════════════════════════════════
// 3. String#[] with range
// ═══════════════════════════════════════════════════════════
#[test] fn str_range_slice() { compile_ok("s = 'hello'\nx = s[1..3]"); }
#[test] fn str_range_exclusive() { compile_ok("s = 'hello'\nx = s[1...4]"); }
#[test] fn arr_range_slice() { compile_ok("a = [1, 2, 3, 4, 5]\nx = a[1..3]"); }

// ═══════════════════════════════════════════════════════════
// 4. for..in with Range
// ═══════════════════════════════════════════════════════════
#[test] fn for_in_range() { compile_ok("for i in 1..10\n  puts i\nend"); }
#[test] fn for_in_exclusive_range() { compile_ok("for i in 0...5\n  puts i\nend"); }

// ═══════════════════════════════════════════════════════════
// 5. Regex capture globals ($1, $2)
// ═══════════════════════════════════════════════════════════
#[test] fn regex_capture_global() { compile_ok("$1 = 'test'\nputs $1"); }
#[test] fn regex_globals() { compile_ok("$0 = 'program'\n$1 = 'match1'"); }

// ═══════════════════════════════════════════════════════════
// 6. String#% format
// ═══════════════════════════════════════════════════════════
#[test] fn str_format_op() { compile_ok("x = 'Hello %s' % 'World'"); }
#[test] fn str_format_number() { compile_ok("x = 'Pi is %.2f' % 3.14"); }

// ═══════════════════════════════════════════════════════════
// 7. Exception hierarchy
// ═══════════════════════════════════════════════════════════
#[test] fn rescue_standard_error() { compile_ok("begin\n  raise 'oops'\nrescue StandardError => e\n  puts e\nend"); }
#[test] fn rescue_runtime_error() { compile_ok("begin\n  raise RuntimeError.new('bad')\nrescue RuntimeError => e\n  puts e\nend"); }

// ═══════════════════════════════════════════════════════════
// 8. Hash.new(default)
// ═══════════════════════════════════════════════════════════
#[test] fn hash_new_default() { compile_ok("h = Hash.new(0)"); }
#[test] fn hash_new_empty() { compile_ok("h = Hash.new"); }

// ═══════════════════════════════════════════════════════════
// 9. Array destructuring in blocks
// ═══════════════════════════════════════════════════════════
#[test] fn block_multi_param() { compile_ok("[[1, 2], [3, 4]].each { |a, b| puts a }"); }
#[test] fn block_three_params() { compile_ok("{a: 1, b: 2}.each_pair { |k, v| puts k }"); }

// ═══════════════════════════════════════════════════════════
// 10. Multi-line method chains
// ═══════════════════════════════════════════════════════════
#[test] fn multiline_chain() { compile_ok("[1, 2, 3]\n  .map { |x| x * 2 }\n  .select { |x| x > 2 }"); }
#[test] fn multiline_chain_dot() { compile_ok("'hello'\n  .upcase\n  .reverse"); }

// ═══════════════════════════════════════════════════════════
// 11. map/select without block (returns self for now)
// ═══════════════════════════════════════════════════════════
#[test] fn map_no_block() { compile_ok("[1, 2, 3].map"); }

// ═══════════════════════════════════════════════════════════
// 12. Real flatten
// ═══════════════════════════════════════════════════════════
#[test] fn flatten_real() { compile_ok("x = [1, 2, 3].flatten"); }

// ═══════════════════════════════════════════════════════════
// 13. Real compact (remove nils)
// ═══════════════════════════════════════════════════════════
#[test] fn compact_real() { compile_ok("x = [1, nil, 2, nil, 3].compact"); }

// ═══════════════════════════════════════════════════════════
// 14. Real uniq (remove duplicates)
// ═══════════════════════════════════════════════════════════
#[test] fn uniq_real() { compile_ok("x = [1, 2, 2, 3, 3].uniq"); }

// ═══════════════════════════════════════════════════════════
// 15. Real zip
// ═══════════════════════════════════════════════════════════
#[test] fn zip_real() { compile_ok("x = [1, 2, 3].zip([4, 5, 6])"); }

// ═══════════════════════════════════════════════════════════
// 16. Real tally
// ═══════════════════════════════════════════════════════════
#[test] fn tally_real() { compile_ok("x = ['a', 'b', 'a', 'c', 'b', 'a'].tally"); }

// ═══════════════════════════════════════════════════════════
// 17. Real group_by
// ═══════════════════════════════════════════════════════════
#[test] fn group_by_real() { compile_ok("x = [1, 2, 3, 4, 5, 6].group_by { |n| n.even? }"); }

// ═══════════════════════════════════════════════════════════
// 18. Struct.new creates real class
// ═══════════════════════════════════════════════════════════
#[test] fn struct_new_class() { compile_ok("Point = Struct.new(:x, :y)"); }

// ═══════════════════════════════════════════════════════════
// 19. Ancestors (returns array)
// ═══════════════════════════════════════════════════════════
#[test] fn ancestors_call() { compile_ok("class Foo\nend\nx = Foo.ancestors"); }

// ═══════════════════════════════════════════════════════════
// 20. Private/protected semantics (parsed, not enforced)
// ═══════════════════════════════════════════════════════════
#[test] fn private_in_class() { compile_ok("class Foo\n  private\n  def secret\n    42\n  end\nend"); }

// ═══════════════════════════════════════════════════════════
// 21. freeze/frozen? semantics
// ═══════════════════════════════════════════════════════════
#[test] fn freeze_returns_self() { compile_ok("x = 'hello'.freeze\nputs x"); }

// ═══════════════════════════════════════════════════════════
// 22. method_missing defined in class
// ═══════════════════════════════════════════════════════════
#[test] fn method_missing_def() { compile_ok("class Ghost\n  def method_missing(name)\n    puts name\n  end\nend\ng = Ghost.new"); }

// ═══════════════════════════════════════════════════════════
// Combined programs using fixed features
// ═══════════════════════════════════════════════════════════
#[test]
fn negative_indexing_program() {
    compile_ok(r#"
arr = [10, 20, 30, 40, 50]
puts arr[-1]
puts arr[-2]
puts arr[0]
"#);
}

#[test]
fn range_iteration() {
    compile_ok(r#"
sum = 0
for i in 1..100
  sum += i
end
puts sum
"#);
}

#[test]
fn string_format_program() {
    compile_ok(r#"
name = 'Ruby'
version = 3.2
msg = 'Welcome to %s' % name
puts msg
"#);
}

#[test]
fn tally_word_count() {
    compile_ok(r#"
words = ['ruby', 'python', 'ruby', 'java', 'python', 'ruby']
counts = words.tally
puts counts.keys.join(', ')
"#);
}

#[test]
fn group_by_parity() {
    compile_ok(r#"
nums = [1, 2, 3, 4, 5, 6, 7, 8]
grouped = nums.group_by { |n| n.even? }
puts grouped.keys.join(', ')
"#);
}

#[test]
fn compact_and_uniq() {
    compile_ok(r#"
data = [1, nil, 2, nil, 3, 2, 1]
clean = data.compact
unique = clean.uniq
puts unique.join(', ')
"#);
}

#[test]
fn multiline_fluent_api() {
    compile_ok(r#"
result = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
  .select { |n| n.even? }
  .map { |n| n * n }
  .sort
puts result.join(', ')
"#);
}

#[test]
fn hash_default_value() {
    compile_ok(r#"
counter = Hash.new(0)
words = ['hello', 'world', 'hello']
puts counter.keys.length
"#);
}

#[test]
fn exception_hierarchy_rescue() {
    compile_ok(r#"
begin
  raise RuntimeError.new('test error')
rescue StandardError => e
  puts 'caught'
rescue => e
  puts 'fallback'
end
"#);
}

#[test]
fn slice_with_range() {
    compile_ok(r#"
arr = [10, 20, 30, 40, 50]
sub = arr[1..3]
puts sub.join(', ')
"#);
}
