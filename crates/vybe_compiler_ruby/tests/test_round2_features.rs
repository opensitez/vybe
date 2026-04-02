use vybe_parser_ruby::parse;
use vybe_compiler_ruby::Compiler;

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
// HIGH IMPACT 1: if/unless/begin as expression
// ═══════════════════════════════════════════════════════════
#[test] fn if_as_expr() { compile_ok("x = if true then 1 else 2 end"); }
#[test] fn if_as_expr_multiline() { compile_ok("x = if true\n  42\nelse\n  0\nend"); }
#[test] fn if_as_expr_elsif() { compile_ok("x = if false\n  1\nelsif true\n  2\nelse\n  3\nend"); }
#[test] fn unless_as_expr() { compile_ok("x = unless false\n  'yes'\nelse\n  'no'\nend"); }
#[test] fn begin_as_expr() { compile_ok("x = begin\n  42\nrescue\n  0\nend"); }
#[test] fn begin_ensure_expr() { compile_ok("x = begin\n  risky_op = 1\nrescue => e\n  0\nensure\n  puts 'done'\nend"); }

// ═══════════════════════════════════════════════════════════
// HIGH IMPACT 2: Safe navigation operator &.
// ═══════════════════════════════════════════════════════════
#[test] fn safe_nav_basic() { compile_ok("x = nil\ny = x&.upcase"); }
#[test] fn safe_nav_chain() { compile_ok("x = nil\ny = x&.strip&.upcase"); }
#[test] fn safe_nav_with_args() { compile_ok("x = 'hello'\ny = x&.gsub('l', 'r')"); }

// ═══════════════════════════════════════════════════════════
// HIGH IMPACT 3: Inline rescue
// ═══════════════════════════════════════════════════════════
#[test] fn inline_rescue() { compile_ok("x = Integer('abc') rescue 0"); }
#[test] fn inline_rescue_method() { compile_ok("result = some_method rescue 'default'"); }

// ═══════════════════════════════════════════════════════════
// HIGH IMPACT 4: Magic constants
// ═══════════════════════════════════════════════════════════
#[test] fn magic_file() { compile_ok("puts __FILE__"); }
#[test] fn magic_line() { compile_ok("puts __LINE__"); }
#[test] fn magic_dir() { compile_ok("puts __dir__"); }
#[test] fn magic_method() { compile_ok("puts __method__"); }

// ═══════════════════════════════════════════════════════════
// HIGH IMPACT 5: Backtick shell commands
// ═══════════════════════════════════════════════════════════
#[test] fn backtick_cmd() { compile_ok("output = `ls -la`"); }
#[test] fn backtick_simple() { compile_ok("x = `echo hello`"); }

// ═══════════════════════════════════════════════════════════
// HIGH IMPACT 6: << operator (string/array append)
// ═══════════════════════════════════════════════════════════
#[test] fn str_append() { compile_ok("s = 'hello'\ns << ' world'"); }
#[test] fn arr_append_op() { compile_ok("a = [1, 2]\na << 3"); }

// ═══════════════════════════════════════════════════════════
// HIGH IMPACT 7: pp (pretty print)
// ═══════════════════════════════════════════════════════════
#[test] fn pp_call() { compile_ok("pp [1, 2, 3]"); }
#[test] fn pp_multi() { compile_ok("pp 'hello', 42"); }
#[test] fn pp_func() { compile_ok("pp(42)"); }

// ═══════════════════════════════════════════════════════════
// HIGH IMPACT 8: sprintf / format
// ═══════════════════════════════════════════════════════════
#[test] fn sprintf_call() { compile_ok("x = sprintf('hello %s', 'world')"); }
#[test] fn format_call() { compile_ok("x = format('%.2f', 3.14)"); }

// ═══════════════════════════════════════════════════════════
// HIGH IMPACT 9: dig (nested access)
// ═══════════════════════════════════════════════════════════
#[test] fn hash_dig() { compile_ok("h = {a: {b: {c: 42}}}\nx = h.dig(:a, :b, :c)"); }
#[test] fn array_dig() { compile_ok("a = [[1, 2], [3, 4]]\nx = a.dig(0, 1)"); }

// ═══════════════════════════════════════════════════════════
// MEDIUM IMPACT 10: filter_map
// ═══════════════════════════════════════════════════════════
#[test] fn filter_map_block() { compile_ok("[1, 2, 3, 4, 5].filter_map { |x| x * 2 if x.odd? }"); }

// ═══════════════════════════════════════════════════════════
// MEDIUM IMPACT 11: tally
// ═══════════════════════════════════════════════════════════
#[test] fn tally_call() { compile_ok("x = ['a', 'b', 'a'].tally"); }

// ═══════════════════════════════════════════════════════════
// MEDIUM IMPACT 12: each_with_object
// ═══════════════════════════════════════════════════════════
#[test] fn each_with_object_block() { compile_ok("[1, 2, 3].each_with_object([]) { |x, arr| arr.push(x * 2) }"); }

// ═══════════════════════════════════════════════════════════
// MEDIUM IMPACT 13: sum with block
// ═══════════════════════════════════════════════════════════
#[test] fn sum_basic() { compile_ok("x = [1, 2, 3].sum"); }

// ═══════════════════════════════════════════════════════════
// MEDIUM IMPACT 14: minmax
// ═══════════════════════════════════════════════════════════
#[test] fn minmax_call() { compile_ok("x = [3, 1, 5, 2].minmax"); }

// ═══════════════════════════════════════════════════════════
// MEDIUM IMPACT 15: Array transforms
// ═══════════════════════════════════════════════════════════
#[test] fn arr_rotate() { compile_ok("x = [1, 2, 3].rotate"); }
#[test] fn arr_transpose() { compile_ok("x = [[1, 2], [3, 4]].transpose"); }
#[test] fn arr_combination() { compile_ok("x = [1, 2, 3].combination(2)"); }

// ═══════════════════════════════════════════════════════════
// MEDIUM IMPACT 16: Integer methods
// ═══════════════════════════════════════════════════════════
#[test] fn int_divmod() { compile_ok("x = 17.divmod(5)"); }
#[test] fn int_digits() { compile_ok("x = 123.digits"); }
#[test] fn int_chr() { compile_ok("x = 65.chr"); }
#[test] fn int_ord() { compile_ok("x = 'A'.ord"); }
#[test] fn int_hex() { compile_ok("x = 255.hex"); }

// ═══════════════════════════════════════════════════════════
// MEDIUM IMPACT 17: IO/STDIN
// ═══════════════════════════════════════════════════════════
#[test] fn warn_call() { compile_ok("warn 'this is a warning'"); }

// ═══════════════════════════════════════════════════════════
// LOWER IMPACT 18: instance_variable_get/set
// ═══════════════════════════════════════════════════════════
#[test] fn ivar_get() { compile_ok("class Foo\n  def initialize\n    @x = 42\n  end\n  def get_x\n    self.instance_variable_get(:x)\n  end\nend"); }
#[test] fn ivar_set() { compile_ok("class Foo\n  def set_x(v)\n    self.instance_variable_set(:x, v)\n  end\nend"); }

// ═══════════════════════════════════════════════════════════
// LOWER IMPACT 19: eql? / equal?
// ═══════════════════════════════════════════════════════════
#[test] fn eql_check() { compile_ok("x = 1.eql?(1)"); }
#[test] fn equal_check() { compile_ok("x = 'a'.equal?('a')"); }

// ═══════════════════════════════════════════════════════════
// LOWER IMPACT 20: encoding
// ═══════════════════════════════════════════════════════════
#[test] fn str_encoding() { compile_ok("x = 'hello'.encoding"); }
#[test] fn str_valid_encoding() { compile_ok("x = 'hello'.valid_encoding?"); }

// ═══════════════════════════════════════════════════════════
// LOWER IMPACT 21: lazy
// ═══════════════════════════════════════════════════════════
#[test] fn lazy_enum() { compile_ok("x = [1, 2, 3].lazy"); }

// ═══════════════════════════════════════════════════════════
// LOWER IMPACT 22: ancestors / class hierarchy
// ═══════════════════════════════════════════════════════════
#[test] fn obj_ancestors() { compile_ok("class Foo\nend\nx = Foo.ancestors"); }

// ═══════════════════════════════════════════════════════════
// LOWER IMPACT 23: hash method
// ═══════════════════════════════════════════════════════════
#[test] fn obj_hash() { compile_ok("x = 'hello'.hash"); }

// ═══════════════════════════════════════════════════════════
// LOWER IMPACT 24: redo
// ═══════════════════════════════════════════════════════════
#[test] fn redo_in_loop() { compile_ok("i = 0\nwhile i < 5\n  i += 1\n  redo if i == 3\nend"); }

// ═══════════════════════════════════════════════════════════
// LOWER IMPACT 25: at_exit
// ═══════════════════════════════════════════════════════════
#[test] fn at_exit_block() { compile_ok("at_exit { puts 'goodbye' }"); }

// ═══════════════════════════════════════════════════════════
// LOWER IMPACT 26: Comparable (via <=>)
// ═══════════════════════════════════════════════════════════
#[test] fn comparable_class() { compile_ok("class Temperature\n  attr_reader :degrees\n  def initialize(d)\n    @degrees = d\n  end\n  def <=>(other)\n    @degrees <=> other.degrees\n  end\nend"); }

// ═══════════════════════════════════════════════════════════
// LOWER IMPACT 27: object introspection
// ═══════════════════════════════════════════════════════════
#[test] fn obj_class_method() { compile_ok("x = 'hello'.class"); }
#[test] fn obj_inspect_method() { compile_ok("x = 42.inspect"); }
#[test] fn obj_dup_method() { compile_ok("a = [1, 2, 3]\nb = a.dup"); }
#[test] fn obj_tap_method() { compile_ok("[1, 2, 3].tap { |arr| puts arr.length }"); }

// ═══════════════════════════════════════════════════════════
// LOWER IMPACT 28: for..in with Range (via array)
// ═══════════════════════════════════════════════════════════
#[test] fn for_in_range() { compile_ok("for i in [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]\n  puts i\nend"); }

// ═══════════════════════════════════════════════════════════
// LOWER IMPACT 29: open (file)
// ═══════════════════════════════════════════════════════════
#[test] fn open_file() { compile_ok("f = open('test.txt', 'r')"); }

// ═══════════════════════════════════════════════════════════
// LOWER IMPACT 30: caller
// ═══════════════════════════════════════════════════════════
#[test] fn caller_call() { compile_ok("x = caller"); }

// ═══════════════════════════════════════════════════════════
// Combined programs using new features
// ═══════════════════════════════════════════════════════════
#[test]
fn safe_nav_real_world() {
    compile_ok(r#"
class User
  attr_accessor :name, :address

  def initialize(name)
    @name = name
    @address = nil
  end
end

user = User.new('Alice')
city = user&.address&.upcase
puts city
"#);
}

#[test]
fn if_expr_assignment() {
    compile_ok(r#"
status = if true
  'active'
else
  'inactive'
end
puts status
"#);
}

#[test]
fn inline_rescue_chain() {
    compile_ok(r#"
data = Integer('not_a_number') rescue 0
puts data
"#);
}

#[test]
fn filter_map_program() {
    compile_ok(r#"
numbers = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
squares_of_odds = numbers.filter_map { |n|
  n * n if n.odd?
}
puts squares_of_odds.join(', ')
"#);
}

#[test]
fn each_with_object_program() {
    compile_ok(r#"
words = ['hello', 'world', 'ruby']
lengths = words.each_with_object({}) { |w, h|
  h[w] = w.length
}
puts lengths.keys.join(', ')
"#);
}

#[test]
fn backtick_program() {
    compile_ok(r#"
output = `echo hello`
puts output
"#);
}

#[test]
fn string_append_program() {
    compile_ok(r#"
greeting = "Hello"
greeting << " "
greeting << "World"
puts greeting
"#);
}

#[test]
fn magic_constants_program() {
    compile_ok(r#"
puts __FILE__
puts __LINE__
puts __dir__
"#);
}

#[test]
fn dig_nested() {
    compile_ok(r#"
config = {database: {host: 'localhost', port: 5432}}
host = config.dig(:database, :host)
puts host
"#);
}

#[test]
fn format_string() {
    compile_ok(r#"
name = 'World'
msg = sprintf('Hello %s!', name)
puts msg
"#);
}

#[test]
fn divmod_program() {
    compile_ok(r#"
result = 17.divmod(5)
puts result.first
puts result.last
"#);
}

#[test]
fn complex_program_with_new_features() {
    compile_ok(r#"
class Config
  attr_reader :settings

  def initialize
    @settings = {}
  end

  def set(key, value)
    @settings[key] = value
  end

  def get(key)
    @settings.dig(key) rescue nil
  end
end

config = Config.new
config.set(:name, 'MyApp')
config.set(:version, '1.0')

name = config&.get(:name)
pp name

status = if config.nil?
  'not loaded'
else
  'loaded'
end
puts status
"#);
}
