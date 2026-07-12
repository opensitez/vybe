//! Python generator compilation via the stack-switching proposal.
//! Functions containing `yield` are compiled with `chunk.is_generator
//! = true`. Calling such a function returns a `Continuation` (lazy);
//! advancing the generator requires explicit RESUME — a primitive the
//! compilers will surface via an iterator protocol in a follow-up.
//!
//! For now these tests verify the binding: yield-containing functions
//! produce generator objects, not eager values.
use super::helpers::run_python;
#[test]
fn generator_decorator_returns_continuation_not_value() {
    // `@generator` opts into true lazy generators via the stack-
    // switching proposal. `gen()` returns a `Continuation` object
    // (displayed as `[continuation]`) — NOT an eager list.
    let out = run_python(
        r#"
@generator
def gen():
    yield 1

x = gen()
print(x)
"#,
    );
    assert_eq!(out, vec!["[continuation]"]);
}

#[test]
fn yield_inside_true_generator_does_not_eagerly_run() {
    // Opt-in via `@generator` — the body must not execute until the
    // continuation is resumed.
    let out = run_python(
        r#"
@generator
def gen():
    print("bad: generator body ran without resume")
    yield 1

_ = gen()
print("ok")
"#,
    );
    assert_eq!(
        out,
        vec!["ok"],
        "@generator body must not run until the continuation is resumed"
    );
}

#[test]
fn for_in_over_true_generator_iterates_lazily() {
    // `@generator` + `for v in gen():` uses the stack-switching
    // iterator protocol. Each loop iteration is a GEN_NEXT: RESUME
    // the continuation, pick up the yielded value, feed it to the
    // body, continue until the generator completes.
    let out = run_python(
        r#"
@generator
def count_to_three():
    yield 1
    yield 2
    yield 3

for v in count_to_three():
    print(v)
"#,
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn true_generator_with_bound_args_yields_once() {
    // Simpler shape: bound arg is passed through and yielded once.
    // Exercises the first-RESUME bound-arg path without dragging in
    // loops inside the generator body.
    let out = run_python(
        r#"
@generator
def ramp(n):
    yield n

for v in ramp(4):
    print(v)
"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn true_generator_with_internal_loop() {
    let out = run_python(
        r#"
@generator
def gen():
    i = 0
    while i < 3:
        yield i
        i = i + 1

for v in gen():
    print(v)
"#,
    );
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn true_generator_with_bound_args_yields_in_order() {
    let out = run_python(
        r#"
@generator
def ramp(n):
    i = 0
    while i < n:
        yield i
        i = i + 1

for v in ramp(4):
    print(v)
"#,
    );
    assert_eq!(out, vec!["0", "1", "2", "3"]);
}

#[test]
fn default_yield_preserves_eager_list_semantics() {
    // Without the `@generator` decorator, yield keeps the legacy
    // eager-list behaviour: `gen()` materialises a list of every
    // yielded value, iterable by `for v in gen()`. This is the
    // backwards-compatible path that existing Python tests rely on.
    let out = run_python(
        r#"
def gen():
    yield 1
    yield 2
    yield 3

for v in gen():
    print(v)
"#,
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}
