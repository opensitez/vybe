use super::helpers::{compile_ok, run_ruby, run_ruby_one};

#[test]
fn custom_exception_class() {
    compile_ok(
        "class MyError < StandardError\nend\n",
    );
}

#[test]
fn custom_exception_custom_message() {
    compile_ok(
        "class AppError < StandardError\n  def initialize(msg = 'app error occurred')\n    super(msg)\n  end\nend\nraise AppError rescue nil\n",
    );
}

#[test]
fn rescue_specific_class() {
    compile_ok(
        "class NetworkError < StandardError\nend\nbegin\n  raise NetworkError\nrescue NetworkError\n  puts 'caught network error'\nend\n",
    );
}

#[test]
fn rescue_capture_exception_object() {
    let out = run_ruby(
        "begin\n  raise 'something went wrong'\nrescue => e\n  puts 'caught'\nend\n",
    );
    assert_eq!(out, vec!["caught"]);
}

#[test]
fn rescue_exception_message() {
    let out = run_ruby(
        "begin\n  raise RuntimeError, 'test message'\nrescue => e\n  puts e.message\nend\n",
    );
    assert_eq!(out, vec!["test message"]);
}

#[test]
fn rescue_exception_class_name() {
    let out = run_ruby(
        "begin\n  raise RuntimeError, 'oops'\nrescue => e\n  puts e.class\nend\n",
    );
    assert_eq!(out, vec!["RuntimeError"]);
}

#[test]
fn multiple_rescue_clauses() {
    compile_ok(
        "class FooError < StandardError\nend\nclass BarError < StandardError\nend\nbegin\n  raise FooError\nrescue FooError\n  puts 'foo'\nrescue BarError\n  puts 'bar'\nend\n",
    );
}

#[test]
fn rescue_comma_separated_types() {
    compile_ok(
        "class FooError < StandardError\nend\nclass BarError < StandardError\nend\nbegin\n  raise FooError\nrescue FooError, BarError\n  puts 'caught one of them'\nend\n",
    );
}

#[test]
fn ensure_runs_on_exception() {
    let out = run_ruby(
        "begin\n  raise 'error'\nrescue\n  puts 'rescued'\nensure\n  puts 'ensured'\nend\n",
    );
    assert_eq!(out, vec!["rescued", "ensured"]);
}

#[test]
fn ensure_runs_on_normal_flow() {
    let out = run_ruby(
        "begin\n  x = 1 + 1\nensure\n  puts 'cleanup'\nend\n",
    );
    assert_eq!(out, vec!["cleanup"]);
}

#[test]
fn retry_inside_rescue() {
    compile_ok(
        "attempts = 0\nbegin\n  attempts += 1\n  raise 'fail' if attempts < 3\nrescue\n  retry if attempts < 3\nend\n",
    );
}

#[test]
fn raise_reraise_current() {
    compile_ok(
        "def risky\n  raise 'original'\nrescue => e\n  raise\nend\nrisky rescue nil\n",
    );
}

#[test]
fn raise_with_new_message() {
    let out = run_ruby(
        "begin\n  raise 'custom message'\nrescue => e\n  puts e.message\nend\n",
    );
    assert_eq!(out, vec!["custom message"]);
}

#[test]
fn raise_explicit_class_and_msg() {
    let out = run_ruby(
        "begin\n  raise RuntimeError, 'explicit msg'\nrescue => e\n  puts e.message\nend\n",
    );
    assert_eq!(out, vec!["explicit msg"]);
}

#[test]
fn begin_rescue_else_no_exception() {
    let out = run_ruby(
        "begin\n  x = 1 + 1\nrescue\n  puts 'error'\nelse\n  puts 'no error'\nend\n",
    );
    assert_eq!(out, vec!["no error"]);
}

#[test]
fn nested_begin_rescue() {
    let out = run_ruby(
        "begin\n  begin\n    raise 'inner'\n  rescue => e\n    puts 'inner caught'\n  end\n  puts 'outer continues'\nrescue\n  puts 'outer caught'\nend\n",
    );
    assert_eq!(out, vec!["inner caught", "outer continues"]);
}

#[test]
fn exception_propagates_from_method() {
    let out = run_ruby(
        "def risky_method\n  raise 'from method'\nend\nbegin\n  risky_method\nrescue => e\n  puts e.message\nend\n",
    );
    assert_eq!(out, vec!["from method"]);
}

#[test]
fn rescue_in_method_body() {
    compile_ok(
        "def safe_divide(a, b)\n  a / b\nrescue ZeroDivisionError\n  0\nend\nsafe_divide(10, 0)\n",
    );
}

#[test]
fn throw_catch_flow_control() {
    compile_ok(
        "result = catch(:done) do\n  [1, 2, 3].each do |n|\n    throw :done, n if n == 2\n  end\nend\n",
    );
}

#[test]
fn catch_returns_throw_value() {
    let out = run_ruby(
        "result = catch(:stop) do\n  throw :stop, 42\nend\nputs result\n",
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn abort_with_message_compiles() {
    compile_ok(
        "def maybe_abort(x)\n  abort('fatal error') if x < 0\n  x\nend\n",
    );
}

#[test]
fn at_exit_register_compiles() {
    compile_ok(
        "at_exit { puts 'cleanup on exit' }\n",
    );
}

#[test]
fn rescue_standard_error_catches_runtime() {
    let out = run_ruby(
        "begin\n  raise RuntimeError, 'boom'\nrescue StandardError => e\n  puts 'caught standard'\nend\n",
    );
    assert_eq!(out, vec!["caught standard"]);
}

#[test]
fn rescue_exception_catches_all() {
    compile_ok(
        "begin\n  raise 'anything'\nrescue Exception\n  puts 'caught all'\nend\n",
    );
}

#[test]
fn custom_exception_with_attributes() {
    compile_ok(
        "class HttpError < StandardError\n  attr_reader :code\n  def initialize(msg, code)\n    super(msg)\n    @code = code\n  end\nend\nbegin\n  raise HttpError.new('not found', 404)\nrescue HttpError => e\n  puts e.code\nend\n",
    );
}

