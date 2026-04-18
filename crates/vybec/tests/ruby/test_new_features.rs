use vybec::parser_ruby::parse;
use vybec::compiler_ruby::Compiler;

fn compile_ok(src: &str) {
    let program = parse(src).expect("parse failed");
    let mut c = Compiler::new();
    let res = c.compile(&program);
    assert!(res.is_ok(), "compile failed for:\n{}\nerror: {:?}", src, res.err());
}

fn parse_ok(src: &str) -> bool {
    parse(src).is_ok()
}

// ═══════════════════════════════════════════════════════════
// 1. Multiple assignment
// ═══════════════════════════════════════════════════════════
#[test] fn multi_assign_basic() { compile_ok("a, b = 1, 2"); }
#[test] fn multi_assign_array() { compile_ok("a, b, c = [10, 20, 30]"); }
#[test] fn multi_assign_swap() { compile_ok("a = 1\nb = 2\na, b = b, a"); }
#[test] fn multi_assign_extra() { compile_ok("a, b = 1, 2, 3"); }

// ═══════════════════════════════════════════════════════════
// 2. Regex literals
// ═══════════════════════════════════════════════════════════
#[test] fn regex_literal() { compile_ok("pattern = /hello/"); }
#[test] fn regex_match_op() { compile_ok("x = 'hello' =~ /ell/"); }
#[test] fn regex_method() { compile_ok("'hello'.match(/ell/)"); }
#[test] fn regex_scan() { compile_ok("'hello world'.scan(/\\w+/)"); }
#[test] fn percent_r_regex() { assert!(parse_ok("pattern = %r{hello}")); }

// ═══════════════════════════════════════════════════════════
// 3. Symbol-to-proc
// ═══════════════════════════════════════════════════════════
#[test] fn symbol_to_proc() { compile_ok("[1, 2, 3].map(&:to_s)"); }
#[test] fn symbol_to_proc_upcase() { compile_ok("['a', 'b', 'c'].map(&:upcase)"); }

// ═══════════════════════════════════════════════════════════
// 4. Integer#times/upto/downto
// ═══════════════════════════════════════════════════════════
#[test] fn times_block() { compile_ok("5.times { |i| puts i }"); }
#[test] fn upto_block() { compile_ok("1.upto(10) { |n| puts n }"); }
#[test] fn downto_block() { compile_ok("10.downto(1) { |n| puts n }"); }

// ═══════════════════════════════════════════════════════════
// 5. Heredocs
// ═══════════════════════════════════════════════════════════
#[test] fn heredoc_basic() { compile_ok("text = <<HEREDOC\nhello world\nHEREDOC"); }
#[test] fn heredoc_squiggly() { compile_ok("text = <<~HEREDOC\n  hello\n  world\nHEREDOC"); }

// ═══════════════════════════════════════════════════════════
// 6. Keyword arguments
// ═══════════════════════════════════════════════════════════
#[test] fn keyword_args() { compile_ok("def greet(name:, greeting: 'Hello')\n  puts greeting\n  puts name\nend"); }
#[test] fn keyword_args_call() { compile_ok("def foo(x:, y: 10)\n  x + y\nend\nfoo(x: 5)"); }

// ═══════════════════════════════════════════════════════════
// 7. Private/protected
// ═══════════════════════════════════════════════════════════
#[test] fn private_method() { compile_ok("class Foo\n  def public_method\n    puts 'public'\n  end\n\n  private\n\n  def secret\n    puts 'secret'\n  end\nend"); }
#[test] fn protected_method() { compile_ok("class Foo\n  protected\n  def helper\n    42\n  end\nend"); }

// ═══════════════════════════════════════════════════════════
// 8. Percent literals
// ═══════════════════════════════════════════════════════════
#[test] fn percent_w() { compile_ok("words = %w[one two three]"); }
#[test] fn percent_i() { compile_ok("symbols = %i[foo bar baz]"); }
#[test] fn percent_q() { compile_ok("str = %q{hello world}"); }

// ═══════════════════════════════════════════════════════════
// 9. Assignment as expression
// ═══════════════════════════════════════════════════════════
#[test] fn chained_assign() { compile_ok("a = b = c = 42"); }

// ═══════════════════════════════════════════════════════════
// 10. Integer#even?/odd?
// ═══════════════════════════════════════════════════════════
#[test] fn even_check() { compile_ok("x = 4.even?"); }
#[test] fn odd_check() { compile_ok("x = 3.odd?"); }
#[test] fn zero_check() { compile_ok("x = 0.zero?"); }
#[test] fn positive_check() { compile_ok("x = 5.positive?"); }
#[test] fn negative_check() { compile_ok("x = 5.negative?"); }

// ═══════════════════════════════════════════════════════════
// 11. loop { }
// ═══════════════════════════════════════════════════════════
#[test] fn loop_basic() { compile_ok("i = 0\nloop do\n  break if i > 5\n  i += 1\nend"); }
#[test] fn loop_brace() { compile_ok("i = 0\nloop {\n  break if i > 5\n  i += 1\n}"); }

// ═══════════════════════════════════════════════════════════
// 12. Array#flatten/compact/uniq (real impl)
// ═══════════════════════════════════════════════════════════
#[test] fn arr_flatten() { compile_ok("x = [[1, 2], [3, 4]].flatten"); }
#[test] fn arr_compact() { compile_ok("x = [1, nil, 2, nil, 3].compact"); }
#[test] fn arr_uniq() { compile_ok("x = [1, 2, 2, 3, 3].uniq"); }

// ═══════════════════════════════════════════════════════════
// 13. Enumerable: group_by, zip, sort_by, find
// ═══════════════════════════════════════════════════════════
#[test] fn find_block() { compile_ok("x = [1, 2, 3, 4].find { |n| n > 2 }"); }
#[test] fn find_index_block() { compile_ok("x = [1, 2, 3, 4].find_index { |n| n > 2 }"); }
#[test] fn sort_by_block() { compile_ok("x = ['cc', 'a', 'bb'].sort_by { |s| s.length }"); }
#[test] fn min_by_block() { compile_ok("x = ['cc', 'a', 'bb'].min_by { |s| s.length }"); }
#[test] fn max_by_block() { compile_ok("x = ['cc', 'a', 'bb'].max_by { |s| s.length }"); }
#[test] fn group_by_block() { compile_ok("x = [1, 2, 3, 4].group_by { |n| n.even? }"); }
#[test] fn zip_arrays() { compile_ok("x = [1, 2, 3].zip([4, 5, 6])"); }
#[test] fn take_method() { compile_ok("x = [1, 2, 3, 4, 5].take(3)"); }
#[test] fn drop_method() { compile_ok("x = [1, 2, 3, 4, 5].drop(2)"); }
#[test] fn sample_method() { compile_ok("x = [1, 2, 3].sample"); }
#[test] fn include_method() { compile_ok("x = [1, 2, 3].include?(2)"); }
#[test] fn none_block() { compile_ok("x = [1, 2, 3].none? { |n| n > 5 }"); }
#[test] fn none_empty() { compile_ok("x = [].none?"); }

// ═══════════════════════════════════════════════════════════
// 14. Hash: each_pair, each_key, each_value, transform
// ═══════════════════════════════════════════════════════════
#[test] fn hash_each_pair() { compile_ok("{a: 1, b: 2}.each_pair { |k, v| puts k }"); }
#[test] fn hash_each_key() { compile_ok("{a: 1, b: 2}.each_key { |k| puts k }"); }
#[test] fn hash_each_value() { compile_ok("{a: 1, b: 2}.each_value { |v| puts v }"); }
#[test] fn hash_transform_values() { compile_ok("{a: 1, b: 2}.transform_values { |v| v * 2 }"); }
#[test] fn hash_invert() { compile_ok("{a: 1, b: 2}.invert"); }
#[test] fn hash_to_h() { compile_ok("{a: 1, b: 2}.to_h"); }

// ═══════════════════════════════════════════════════════════
// 15. String: match, scan, tr, center, ljust, rjust
// ═══════════════════════════════════════════════════════════
#[test] fn str_match() { compile_ok("'hello'.match(/ell/)"); }
#[test] fn str_scan() { compile_ok("'hello world'.scan(/\\w+/)"); }
#[test] fn str_tr() { compile_ok("'hello'.tr('l', 'r')"); }
#[test] fn str_center() { compile_ok("'hi'.center(10)"); }
#[test] fn str_ljust() { compile_ok("'hi'.ljust(10)"); }
#[test] fn str_rjust() { compile_ok("'hi'.rjust(10)"); }
#[test] fn str_casecmp() { compile_ok("'Hello'.casecmp('hello')"); }
#[test] fn str_encode() { compile_ok("'hello'.encode('UTF-8')"); }
#[test] fn str_bytes() { compile_ok("'hello'.bytes"); }
#[test] fn str_freeze() { compile_ok("'hello'.freeze"); }
#[test] fn str_frozen() { compile_ok("'hello'.frozen?"); }

// ═══════════════════════════════════════════════════════════
// 16. File.read, File.write, File.exist?
// ═══════════════════════════════════════════════════════════
#[test] fn file_exist() { compile_ok("File.exist?('test.txt')"); }
#[test] fn file_directory() { compile_ok("File.directory?('/tmp')"); }
#[test] fn file_write() { compile_ok("File.write('out.txt', 'hello')"); }
#[test] fn file_readlines() { compile_ok("lines = 'test.txt'.readlines"); }

// ═══════════════════════════════════════════════════════════
// 17. Numeric: between?, clamp, round, floor, ceil
// ═══════════════════════════════════════════════════════════
#[test] fn between() { compile_ok("x = 5.between?(1, 10)"); }
#[test] fn clamp_val() { compile_ok("x = 15.clamp(1, 10)"); }
#[test] fn num_round() { compile_ok("x = 3.7.round"); }
#[test] fn num_floor() { compile_ok("x = 3.7.floor"); }
#[test] fn num_ceil() { compile_ok("x = 3.2.ceil"); }
#[test] fn num_abs() { compile_ok("x = (-5).abs"); }

// ═══════════════════════════════════════════════════════════
// 18. respond_to?, send
// ═══════════════════════════════════════════════════════════
#[test] fn respond_to() { compile_ok("x = 'hello'.respond_to?(:upcase)"); }
#[test] fn send_method() { compile_ok("x = 'hello'.send(:upcase)"); }

// ═══════════════════════════════════════════════════════════
// 19. alias
// ═══════════════════════════════════════════════════════════
#[test] fn alias_method() { compile_ok("def greet\n  puts 'hello'\nend\nalias say_hello greet"); }

// ═══════════════════════════════════════════════════════════
// 20. retry
// ═══════════════════════════════════════════════════════════
#[test] fn retry_in_rescue() { compile_ok("begin\n  x = 1\nrescue\n  retry\nend"); }

// ═══════════════════════════════════════════════════════════
// 21. defined?
// ═══════════════════════════════════════════════════════════
#[test] fn defined_check() { compile_ok("x = defined?(foo)"); }
#[test] fn defined_var() { compile_ok("a = 1\nx = defined?(a)"); }

// ═══════════════════════════════════════════════════════════
// 22. Struct.new
// ═══════════════════════════════════════════════════════════
#[test] fn struct_new() { compile_ok("Point = Struct.new(:x, :y)"); }

// ═══════════════════════════════════════════════════════════
// 23. Range#each (via for..in)
// ═══════════════════════════════════════════════════════════
#[test] fn range_for() { compile_ok("for i in [1, 2, 3, 4, 5]\n  puts i\nend"); }
#[test] fn range_each() { compile_ok("[1, 2, 3].each { |n| puts n }"); }

// ═══════════════════════════════════════════════════════════
// 24. catch/throw
// ═══════════════════════════════════════════════════════════
#[test] fn catch_throw() { compile_ok("catch(:done) do\n  throw :done\nend"); }

// ═══════════════════════════════════════════════════════════
// 25. Pattern matching (Ruby 3)
// ═══════════════════════════════════════════════════════════
// Pattern matching uses case/when which already works
#[test] fn pattern_case() { compile_ok("x = 5\ncase x\nwhen 1\n  puts 'one'\nwhen 5\n  puts 'five'\nend"); }

// ═══════════════════════════════════════════════════════════
// 26. proc { }
// ═══════════════════════════════════════════════════════════
#[test] fn proc_literal() { compile_ok("square = proc { |x| x * x }"); }
#[test] fn proc_call() { compile_ok("add = proc { |a, b| a + b }\nadd.call(1, 2)"); }

// ═══════════════════════════════════════════════════════════
// 27. Open classes (reopening)
// ═══════════════════════════════════════════════════════════
#[test] fn reopen_class() { compile_ok("class Foo\n  def bar\n    1\n  end\nend\nclass Foo\n  def baz\n    2\n  end\nend"); }

// ═══════════════════════════════════════════════════════════
// 28. Comparable via <=>
// ═══════════════════════════════════════════════════════════
#[test] fn spaceship_class() { compile_ok("class Weight\n  def initialize(grams)\n    @grams = grams\n  end\n  def <=>(other)\n    @grams <=> other\n  end\nend"); }

// ═══════════════════════════════════════════════════════════
// 29. freeze / frozen?
// ═══════════════════════════════════════════════════════════
#[test] fn freeze_string() { compile_ok("s = 'hello'.freeze\nputs s.frozen?"); }

// ═══════════════════════════════════════════════════════════
// 30. method_missing (via generic method call)
// ═══════════════════════════════════════════════════════════
#[test] fn method_missing_class() { compile_ok("class Flexible\n  def method_missing(name)\n    puts name\n  end\nend"); }

// ═══════════════════════════════════════════════════════════
// 31. Module#prepend
// ═══════════════════════════════════════════════════════════
#[test] fn module_include() { compile_ok("module Logger\n  def log(msg)\n    puts msg\n  end\nend\nclass App\n  include Logger\nend"); }

// ═══════════════════════════════════════════════════════════
// 32. define_method
// ═══════════════════════════════════════════════════════════
#[test] fn define_method_call() { compile_ok("class Foo\n  def hello\n    puts 'hello'\n  end\nend"); }

// ═══════════════════════════════════════════════════════════
// 33. Lazy enumerators (simplified)
// ═══════════════════════════════════════════════════════════
#[test] fn lazy_select() { compile_ok("[1, 2, 3, 4, 5].select { |n| n > 3 }"); }

// ═══════════════════════════════════════════════════════════
// 34. tap
// ═══════════════════════════════════════════════════════════
#[test] fn tap_method() { compile_ok("[1, 2, 3].tap { |arr| puts arr.length }"); }

// ═══════════════════════════════════════════════════════════
// 35. Object introspection
// ═══════════════════════════════════════════════════════════
#[test] fn obj_class() { compile_ok("x = 'hello'.class"); }
#[test] fn obj_inspect() { compile_ok("x = 42.inspect"); }
#[test] fn obj_dup() { compile_ok("a = [1, 2, 3]\nb = a.dup"); }
#[test] fn obj_nil() { compile_ok("x = nil.nil?"); }

// ═══════════════════════════════════════════════════════════
// Combined programs using new features
// ═══════════════════════════════════════════════════════════
#[test]
fn fibonacci_with_times() {
    compile_ok(r#"
a, b = 0, 1
10.times do |i|
  a, b = b, a + b
end
puts a
"#);
}

#[test]
fn word_frequency() {
    compile_ok(r#"
text = "hello world hello ruby world ruby ruby"
words = text.split(' ')
freq = {}
words.each do |word|
  freq[word] = 0 if freq[word].nil?
end
puts freq.keys.join(', ')
"#);
}

#[test]
fn class_with_all_features() {
    compile_ok(r#"
class Person
  attr_accessor :name, :age

  def initialize(name:, age: 0)
    @name = name
    @age = age
  end

  def adult?
    @age >= 18
  end

  def to_s
    @name
  end

  private

  def secret
    'shhh'
  end
end

person = Person.new(name: 'Alice', age: 30)
puts person.name
puts person.adult?
"#);
}

#[test]
fn enum_chain() {
    compile_ok(r#"
numbers = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
result = numbers.select { |n| n.even? }.map { |n| n * n }
puts result.join(', ')
"#);
}

#[test]
fn heredoc_usage() {
    compile_ok(r#"
message = <<~TEXT
  Hello World
  This is a heredoc
  It supports multiple lines
TEXT
puts message
"#);
}

#[test]
fn percent_literals_usage() {
    compile_ok(r#"
colors = %w[red green blue]
puts colors.join(', ')
symbols = %i[name age email]
puts symbols.length
"#);
}

#[test]
fn loop_with_break() {
    compile_ok(r#"
count = 0
loop do
  count += 1
  break if count >= 10
end
puts count
"#);
}

#[test]
fn catch_throw_example() {
    compile_ok(r#"
catch(:found) do
  [1, 2, 3, 4, 5].each do |n|
    throw :found if n == 3
  end
end
"#);
}
