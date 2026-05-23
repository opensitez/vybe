use super::helpers::compile_ok;

// ── Basic case/in pattern matching (Ruby 3+) ──────────────────
#[test] fn case_in_integer_match() { compile_ok(r#"
value = 42
case value
in Integer => n
  puts "integer: #{n}"
end
"#); }

#[test] fn case_in_string_match() { compile_ok(r#"
val = "hello"
case val
in String => s
  puts "string: #{s}"
end
"#); }

// ── Array pattern ─────────────────────────────────────────────
#[test] fn array_pattern_first_rest() { compile_ok(r#"
case [1, 2, 3]
in [first, *rest]
  puts first
  puts rest.inspect
end
"#); }

#[test] fn array_pattern_fixed_length() { compile_ok(r#"
case [10, 20]
in [a, b]
  puts a + b
end
"#); }

#[test] fn array_pattern_nested() { compile_ok(r#"
case [[1, 2], [3, 4]]
in [[a, b], [c, d]]
  puts a + b + c + d
end
"#); }

// ── Hash pattern ──────────────────────────────────────────────
#[test] fn hash_pattern_extracts_keys() { compile_ok(r#"
data = { name: "Alice", age: 30 }
case data
in { name: String => name, age: Integer => age }
  puts name.to_s + ' is ' + age.to_s
end
"#); }

#[test] fn hash_pattern_partial_match() { compile_ok(r#"
event = { type: :click, x: 100, y: 200 }
case event
in { type: :click, x: Integer => x }
  puts 'click at x=' + x.to_s
end
"#); }

// ── Find pattern ─────────────────────────────────────────────
#[test] fn find_pattern_in_array() { compile_ok(r#"
case [1, 2, 42, 3, 4]
in [*, 42, *]
  puts "found 42"
end
"#); }

// ── Pin operator ─────────────────────────────────────────────
#[test] fn pin_operator_matches_variable() { compile_ok(r#"
expected = 42
case [1, 42, 3]
in [*, ^expected, *]
  puts "found expected"
end
"#); }

// ── Guard clauses ─────────────────────────────────────────────
#[test] fn pattern_with_if_guard() { compile_ok(r#"
case 15
in n if n > 10
  puts "big: #{n}"
end
"#); }

// ── Deconstruct protocol ─────────────────────────────────────
#[test] fn custom_deconstruct_for_array_pattern() { compile_ok(r#"
class Point
  attr_reader :x, :y
  def initialize(x, y); @x = x; @y = y; end
  def deconstruct; [@x, @y]; end
end
case Point.new(3, 4)
in [x, y]
  puts x + y
end
"#); }

#[test] fn custom_deconstruct_keys_for_hash_pattern() { compile_ok(r#"
class Config
  def initialize(h, p); @host = h; @port = p; end
  def deconstruct_keys(keys); { host: @host, port: @port }; end
end
case Config.new("localhost", 8080)
in { host: String => h, port: Integer => p }
  puts h.to_s + ':' + p.to_s
end
"#); }

// ── Multiple patterns with | ───────────────────────────────────
#[test] fn pattern_or_alternatives() { compile_ok(r#"
[1, 2, 3].each do |n|
  case n
  in 1 | 3
    puts "odd"
  in 2
    puts "even"
  end
end
"#); }

// ── Nested hash with array ────────────────────────────────────
#[test] fn pattern_nested_hash_and_array() { compile_ok(r#"
response = { status: 200, body: ["ok", "done"] }
case response
in { status: 200, body: [first, *] }
  puts first
end
"#); }

// ── in pattern as expression (one-liner) ──────────────────────
#[test] fn pattern_match_returns_bool() { compile_ok(r#"
result = [1, 2, 3] in [Integer, Integer, Integer]
puts result
"#); }

// ── else branch ───────────────────────────────────────────────
#[test] fn pattern_with_else_fallthrough() { compile_ok(r#"
case { type: :unknown }
in { type: :click }
  puts "click"
in { type: :keypress }
  puts "keypress"
else
  puts "other"
end
"#); }

// ── Rightward assignment pattern ──────────────────────────────
#[test] fn rightward_assignment_stores_match() { compile_ok(r#"
{ name: "Bob", score: 95 } => { name: String => player, score: Integer => pts }
puts player
puts pts
"#); }

// ── Range in pattern ──────────────────────────────────────────
#[test] fn pattern_range_matching() { compile_ok(r#"
score = 85
case score
in 90..100
  puts "A"
in 80...90
  puts "B"
in 70...80
  puts "C"
else
  puts "F"
end
"#); }
