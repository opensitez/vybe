use super::helpers::{compile_ok, run_ruby, run_ruby_one};

// ── block_given? ─────────────────────────────────────────────────────────────

#[test]
fn block_given_check() {
    compile_ok(
        "def maybe_yield\n  if block_given?\n    yield\n  else\n    puts 'no block'\n  end\nend\n",
    );
}

// ── block yield return value ──────────────────────────────────────────────────

#[test]
fn block_yield_return_value_used() {
    compile_ok(
        "def transform(x)\n  yield x\nend\nresult = transform(5) { |n| n * 3 }\n",
    );
}

// ── block closing over outer variable ────────────────────────────────────────

#[test]
fn block_closure_captures_outer_var() {
    compile_ok("x = 10\nf = proc { puts x }\n");
}

// ── block modifying outer variable ───────────────────────────────────────────

#[test]
fn block_closure_modifies_outer_var() {
    compile_ok("total = 0\n[1, 2, 3].each { |n| total += n }\n");
}

// ── explicit block parameter &block ──────────────────────────────────────────

#[test]
fn explicit_block_parameter() {
    compile_ok("def run(&block)\n  block.call\nend\n");
}

// ── calling explicit block with block.call ────────────────────────────────────

#[test]
fn explicit_block_call() {
    compile_ok("def run(&block)\n  block.call(42)\nend\nrun { |x| puts x }\n");
}

// ── forwarding &block to another method ──────────────────────────────────────

#[test]
fn block_forwarding_to_another_method() {
    compile_ok(
        "def outer(&block)\n  inner(&block)\nend\ndef inner\n  yield 1\nend\n",
    );
}

// ── Proc.new from explicit block ──────────────────────────────────────────────

#[test]
fn proc_new_from_block() {
    compile_ok("p = Proc.new { |x| x + 1 }\n");
}

// ── proc { } shorthand ───────────────────────────────────────────────────────

#[test]
fn proc_shorthand_syntax() {
    compile_ok("p = proc { |x| x * 2 }\n");
}

// ── lambda { } long syntax ───────────────────────────────────────────────────

#[test]
fn lambda_long_syntax() {
    compile_ok("f = lambda { |x| x + 1 }\n");
}

// ── stabby lambda with multiple params ───────────────────────────────────────

#[test]
fn lambda_stabby_multi_param() {
    compile_ok("add = ->(a, b) { a + b }\n");
}

// ── lambda vs proc return behavior ───────────────────────────────────────────

#[test]
fn lambda_vs_proc_return_behavior() {
    compile_ok(
        "def test_lambda\n  f = lambda { return 1 }\n  f.call\n  2\nend\n",
    );
}

// ── lambda arity enforcement ──────────────────────────────────────────────────

#[test]
fn lambda_arity_declaration() {
    compile_ok("f = ->(a, b, c) { a + b + c }\n");
}

// ── curry partial application ─────────────────────────────────────────────────

#[test]
fn curry_partial_application() {
    compile_ok("add = ->(a, b) { a + b }\ncurried = add.curry\n");
}

// ── curry calling with remaining args ─────────────────────────────────────────

#[test]
fn curry_apply_remaining_args() {
    compile_ok(
        "add = ->(a, b) { a + b }\nadd5 = add.curry.(5)\nresult = add5.(3)\n",
    );
}

// ── arity of proc/lambda ──────────────────────────────────────────────────────

#[test]
fn proc_arity_method() {
    compile_ok("f = ->(a, b) { a + b }\nn = f.arity\n");
}

// ── call via .call, .() and [] ────────────────────────────────────────────────

#[test]
fn proc_call_three_syntaxes() {
    compile_ok(
        "f = ->(x) { x * 2 }\na = f.call(3)\nb = f.(3)\nc = f[3]\n",
    );
}

// ── &method(:name) — method reference as proc ────────────────────────────────

#[test]
fn method_ref_as_proc() {
    compile_ok(
        "def double(x)\n  x * 2\nend\nresult = [1, 2, 3].map(&method(:double))\n",
    );
}

// ── &:symbol — symbol to proc ────────────────────────────────────────────────

#[test]
fn symbol_to_proc() {
    compile_ok("result = ['hello', 'world'].map(&:upcase)\n");
}

// ── tap — chain with side effect ──────────────────────────────────────────────

#[test]
fn tap_chain_side_effect() {
    compile_ok(
        "result = [1, 2, 3].tap { |a| puts a.length }.map { |x| x * 2 }\n",
    );
}

// ── then / yield_self — pipe value through block ─────────────────────────────

#[test]
fn then_pipes_value_through_block() {
    compile_ok("result = 5.then { |x| x * 2 }\n");
}

// ── itself — returns receiver unchanged ──────────────────────────────────────

#[test]
fn itself_returns_receiver() {
    compile_ok("x = 42.itself\n");
}

// ── closure retaining value after outer scope ends ───────────────────────────

#[test]
fn closure_retains_value_after_scope() {
    compile_ok(
        "def make_adder(n)\n  ->(x) { x + n }\nend\nadd10 = make_adder(10)\n",
    );
}

// ── multiple yield calls in one method ───────────────────────────────────────

#[test]
fn multiple_yield_calls() {
    compile_ok(
        "def three_times\n  yield 1\n  yield 2\n  yield 3\nend\n",
    );
}

// ── yield with multiple values ────────────────────────────────────────────────

#[test]
fn yield_multiple_values() {
    compile_ok("def pair\n  yield 'key', 'value'\nend\n");
}

// ── block with next to skip iteration value ───────────────────────────────────

#[test]
fn block_next_skip_value() {
    compile_ok(
        "result = [1, 2, 3, 4].map { |x| next 0 if x.even?; x }\n",
    );
}

// ── block with break to exit and return value ─────────────────────────────────

#[test]
fn block_break_with_return_value() {
    compile_ok(
        "result = [1, 2, 3, 4].each { |x| break x if x > 2 }\n",
    );
}

// ── each_with_object ──────────────────────────────────────────────────────────

#[test]
fn each_with_object_accumulator() {
    compile_ok(
        "result = [1, 2, 3].each_with_object([]) { |x, acc| acc.push(x * 2) }\n",
    );
}

// ── recursive proc via variable capture ──────────────────────────────────────

#[test]
fn recursive_proc_via_capture() {
    compile_ok(
        "fib = nil\nfib = ->(n) { n < 2 ? n : fib.(n - 1) + fib.(n - 2) }\n",
    );
}

// ── memoization using closure ─────────────────────────────────────────────────

#[test]
fn memoization_via_closure() {
    compile_ok(
        "cache = {}\nmemo = ->(n) { cache[n] ||= n * n }\n",
    );
}

// ── lazy enumerator ───────────────────────────────────────────────────────────

#[test]
fn lazy_enumerator_creation() {
    compile_ok("e = (1..Float::INFINITY).lazy\n");
}

// ── lazy.map.first(n) ─────────────────────────────────────────────────────────

#[test]
fn lazy_map_first_n() {
    compile_ok(
        "result = (1..Float::INFINITY).lazy.map { |x| x * 2 }.first(5)\n",
    );
}

// ── Enumerator.new custom enumerator ─────────────────────────────────────────

#[test]
fn enumerator_new_custom() {
    compile_ok(
        "e = Enumerator.new { |y| y << 1; y << 2; y << 3 }\n",
    );
}

// ── Enumerator::Lazy chained operations ───────────────────────────────────────

#[test]
fn enumerator_lazy_chained() {
    compile_ok(
        "result = [1, 2, 3, 4, 5].lazy.select { |x| x.odd? }.map { |x| x * 10 }.first(2)\n",
    );
}

// ── Object#freeze and block side effects ──────────────────────────────────────

#[test]
fn freeze_object_then_use_in_block() {
    compile_ok(
        "s = 'hello'.freeze\nresult = [s].map { |x| x.length }\n",
    );
}

// ── Proc composition with >> ──────────────────────────────────────────────────

#[test]
fn proc_compose_right() {
    compile_ok(
        "double = ->(x) { x * 2 }\nincrement = ->(x) { x + 1 }\ndouble_then_inc = double >> increment\n",
    );
}

// ── Proc composition with << ──────────────────────────────────────────────────

#[test]
fn proc_compose_left() {
    compile_ok(
        "double = ->(x) { x * 2 }\nincrement = ->(x) { x + 1 }\ninc_then_double = double << increment\n",
    );
}

// ── Method object calling with .() ────────────────────────────────────────────

#[test]
fn method_object_call_syntax() {
    compile_ok(
        "def square(x)\n  x * x\nend\nm = method(:square)\nresult = m.(5)\n",
    );
}

// ── respond_to? checking for method existence ────────────────────────────────

#[test]
fn respond_to_method_check() {
    compile_ok("x = 'hello'\nputs x.respond_to?(:upcase)\n");
}

// ── send calling method by name ───────────────────────────────────────────────

#[test]
fn send_call_by_name() {
    compile_ok("result = 'hello'.send(:upcase)\n");
}

// ── runtime: block_given? false path ─────────────────────────────────────────

#[test]
fn block_given_false_runtime() {
    let out = run_ruby(
        "def maybe_yield\n  if block_given?\n    yield\n  else\n    puts 'no block'\n  end\nend\nmaybe_yield\n",
    );
    assert_eq!(out, vec!["no block"]);
}

// ── runtime: proc shorthand executes ─────────────────────────────────────────

#[test]
fn proc_shorthand_runtime() {
    assert_eq!(
        run_ruby_one("p = proc { |x| x * 3 }\nputs p.call(4)\n"),
        "12"
    );
}

// ── runtime: closure captures outer value ────────────────────────────────────

#[test]
fn closure_captures_value_runtime() {
    assert_eq!(
        run_ruby_one("n = 7\nf = -> { n * 2 }\nputs f.()\n"),
        "14"
    );
}

// ── runtime: each_with_object builds array ───────────────────────────────────

#[test]
fn each_with_object_runtime() {
    let out = run_ruby(
        "r = [1, 2, 3].each_with_object([]) { |x, acc| acc.push(x * 10) }\nr.each { |v| puts v }\n",
    );
    assert_eq!(out, vec!["10", "20", "30"]);
}
