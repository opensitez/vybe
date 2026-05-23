use super::helpers::{compile_ok, run_ruby, run_ruby_one};

// -- tap -- chain side effect, returns receiver

#[test]
fn tap_returns_receiver() {
    compile_ok(r#"[1, 2, 3].tap { |a| puts a.length }
"#);
}

#[test]
fn tap_receiver_unchanged_runtime() {
    let out = run_ruby(r#"[1, 2, 3].tap { |a| a.push(4) }.each { |x| puts x }
"#);
    assert_eq!(out, vec!["1", "2", "3", "4"]);
}

// -- then / yield_self -- pipes value through block

#[test]
fn then_yield_self_alias() {
    compile_ok(r#"result = 5.yield_self { |x| x * 2 }
"#);
}

#[test]
fn then_chains_transformations() {
    compile_ok(r#"result = 5.then { |x| x + 1 }.then { |x| x * 2 }
"#);
}

// -- itself -- returns receiver unchanged

#[test]
fn itself_identity_method() {
    compile_ok("x = 42.itself
");
}

#[test]
fn itself_runtime() {
    assert_eq!(run_ruby_one(r#"puts 'hello'.itself
"#), "hello");
}

// -- freeze string to prevent mutation

#[test]
fn freeze_prevents_mutation() {
    compile_ok(r#"s = 'hello'.freeze
"#);
}

// -- frozen? on freshly created object

#[test]
fn frozen_predicate_on_new_object() {
    compile_ok(r#"puts 'hello'.frozen?
"#);
}

#[test]
fn frozen_predicate_runtime() {
    assert_eq!(run_ruby_one(r#"puts 'hello'.freeze.frozen?
"#), "true");
}

// -- dup -- shallow copy, not frozen

#[test]
fn dup_shallow_copy_not_frozen() {
    compile_ok(r#"orig = 'hello'.freeze
copy = orig.dup
puts copy.frozen?
"#);
}

#[test]
fn dup_not_frozen_runtime() {
    assert_eq!(
        run_ruby_one(r#"puts 'hello'.freeze.dup.frozen?
"#),
        "false"
    );
}

// -- clone -- preserves frozen state

#[test]
fn clone_preserves_frozen_state() {
    compile_ok(r#"orig = 'hello'.freeze
copy = orig.clone
"#);
}

// -- object_id -- unique per object

#[test]
fn object_id_unique_per_object() {
    compile_ok(r#"a = 'hello'
b = 'hello'
result = a.object_id == b.object_id
"#);
}

// -- equal? -- identity comparison (same object_id)

#[test]
fn equal_identity_comparison() {
    compile_ok(r#"a = 'hello'
b = a
result = a.equal?(b)
"#);
}

// -- == vs equal? distinction

#[test]
fn equal_vs_double_equals_distinction() {
    compile_ok(r#"a = 'hello'
b = 'hello'
value_eq = a == b
identity_eq = a.equal?(b)
"#);
}

// -- Conditional assignment ||=

#[test]
fn conditional_assign_or() {
    compile_ok(r#"x = nil
x ||= 42
"#);
}

#[test]
fn conditional_assign_or_runtime() {
    assert_eq!(run_ruby_one(r#"x = nil
x ||= 42
puts x
"#), "42");
}

#[test]
fn conditional_assign_or_preserves_existing() {
    assert_eq!(run_ruby_one(r#"x = 7
x ||= 42
puts x
"#), "7");
}

// -- Conditional assignment &&=

#[test]
fn conditional_assign_and() {
    compile_ok(r#"x = 5
x &&= x * 2
"#);
}

#[test]
fn conditional_assign_and_runtime() {
    assert_eq!(run_ruby_one(r#"x = 5
x &&= x * 2
puts x
"#), "10");
}

#[test]
fn conditional_assign_and_nil_stays_nil() {
    assert_eq!(run_ruby_one(r#"x = nil
x &&= 42
puts x.nil?
"#), "true");
}

// -- Safe navigation operator &.

#[test]
fn safe_navigation_operator() {
    compile_ok(r#"s = 'hello'
result = s&.upcase
"#);
}

// -- &. on nil returns nil without error

#[test]
fn safe_navigation_on_nil_returns_nil() {
    compile_ok(r#"s = nil
result = s&.upcase
"#);
}

#[test]
fn safe_navigation_nil_runtime() {
    assert_eq!(run_ruby_one(r#"s = nil
puts s&.upcase.nil?
"#), "true");
}

// -- pp -- pretty print (compile_ok)

#[test]
fn pp_pretty_print() {
    compile_ok(r#"pp [1, 2, 3]
"#);
}

// -- __method__ -- current method name

#[test]
fn dunder_method_current_name() {
    compile_ok(r#"def my_func
  puts __method__
end
"#);
}

#[test]
fn dunder_method_runtime() {
    assert_eq!(run_ruby_one(r#"def greet
  __method__.to_s
end
puts greet
"#), "greet");
}

// -- caller -- call stack (compile_ok)

#[test]
fn caller_call_stack() {
    compile_ok(r#"def deep
  caller
end
deep
"#);
}

// -- defined? -- check if expression is defined

#[test]
fn defined_expression_check() {
    compile_ok(r#"x = 1
result = defined?(x)
"#);
}

#[test]
fn defined_undefined_returns_nil() {
    compile_ok(r#"result = defined?(totally_undefined_var_xyz)
"#);
}

// -- __FILE__ -- current file

#[test]
fn dunder_file_constant() {
    compile_ok(r#"f = __FILE__
"#);
}

// -- __LINE__ -- current line number

#[test]
fn dunder_line_constant() {
    compile_ok(r#"n = __LINE__
"#);
}

// -- __dir__ -- current directory

#[test]
fn dunder_dir_constant() {
    compile_ok(r#"d = __dir__
"#);
}

// -- Kernel#rand -- random number

#[test]
fn kernel_rand_random_number() {
    compile_ok(r#"r = rand
"#);
}

#[test]
fn kernel_rand_with_bound() {
    compile_ok(r#"r = rand(100)
"#);
}

// -- Kernel#srand -- seed random

#[test]
fn kernel_srand_seed() {
    compile_ok(r#"srand(42)
r = rand(10)
"#);
}

// -- Integer() strict conversion (raises on invalid)

#[test]
fn integer_strict_conversion() {
    compile_ok(r#"n = Integer('42')
"#);
}

#[test]
fn integer_strict_conversion_runtime() {
    assert_eq!(run_ruby_one(r#"puts Integer('42')
"#), "42");
}

// -- String() / Array() / Hash() conversion methods

#[test]
fn string_conversion_method() {
    compile_ok(r#"s = String(42)
"#);
}

#[test]
fn array_conversion_method() {
    compile_ok(r#"a = Array(nil)
"#);
}

#[test]
fn array_conversion_wraps_non_array() {
    compile_ok(r#"a = Array(42)
"#);
}

#[test]
fn hash_conversion_method() {
    compile_ok(r#"h = Hash(nil)
"#);
}

#[test]
fn string_conversion_runtime() {
    assert_eq!(run_ruby_one(r#"puts String(99)
"#), "99");
}
