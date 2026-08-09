/// Generator advanced — send values via next(), throw(), return(), delegation,
/// generator as state machine, lazy evaluation, producer/consumer pattern.
use super::helpers::run_js;

// ── sending values to generators ──────────────────────────────────────────────

#[test]
fn generator_next_with_value_received_at_yield() {
    assert_eq!(
        run_js(
            r#"
function* dialog() {
    const name = yield "What's your name?";
    const age = yield `Hello ${name}, how old are you?`;
    yield `${name} is ${age} years old`;
}
const g = dialog();
console.log(g.next().value);       // "What's your name?"
console.log(g.next("Alice").value); // "Hello Alice, how old are you?"
console.log(g.next(30).value);      // "Alice is 30 years old"
"#
        ),
        vec![
            "What's your name?",
            "Hello Alice, how old are you?",
            "Alice is 30 years old"
        ]
    );
}

#[test]
fn generator_first_next_arg_is_always_ignored() {
    assert_eq!(
        run_js(
            r#"
function* gen() {
    const x = yield 1;
    yield x;
}
const g = gen();
g.next("ignored"); // first next arg always ignored
const r = g.next(42);
console.log(r.value); // 42 — second next value received
"#
        ),
        vec!["42"]
    );
}

// ── generator throw ───────────────────────────────────────────────────────────

#[test]
fn generator_throw_causes_error_at_yield_point() {
    assert_eq!(
        run_js(
            r#"
function* gen() {
    try {
        yield 1;
    } catch (e) {
        yield "caught:" + e.message;
    }
}
const g = gen();
g.next(); // advance to yield 1
const result = g.throw(new Error("boom"));
console.log(result.value);
console.log(result.done);
"#
        ),
        vec!["caught:boom", "false"]
    );
}

#[test]
fn generator_throw_propagates_if_uncaught() {
    assert_eq!(
        run_js(
            r#"
function* gen() {
    yield 1;
    yield 2;
}
const g = gen();
g.next();
let threw = false;
try {
    g.throw(new Error("err"));
} catch (e) {
    threw = true;
}
console.log(threw);
"#
        ),
        vec!["true"]
    );
}

// ── generator return ──────────────────────────────────────────────────────────

#[test]
fn generator_return_ends_iteration() {
    assert_eq!(
        run_js(
            r#"
function* count() {
    yield 1; yield 2; yield 3;
}
const g = count();
g.next(); // 1
const r = g.return("done");
console.log(r.value);
console.log(r.done);
const next = g.next();
console.log(next.done); // still done
"#
        ),
        vec!["done", "true", "true"]
    );
}

#[test]
fn generator_return_triggers_finally() {
    assert_eq!(
        run_js(
            r#"
function* gen() {
    try {
        yield 1;
        yield 2;
    } finally {
        yield "cleanup";
    }
}
const g = gen();
g.next(); // advance to yield 1
const r = g.return("early");
// return causes finally to run, yielding "cleanup"
console.log(r.value);
console.log(r.done);
"#
        ),
        vec!["cleanup", "false"]
    );
}

// ── yield* delegation ─────────────────────────────────────────────────────────

#[test]
fn yield_star_delegates_completely() {
    assert_eq!(
        run_js(
            r#"
function* inner() { yield "a"; yield "b"; }
function* outer() { yield* inner(); yield "c"; }
console.log([...outer()].join(","));
"#
        ),
        vec!["a,b,c"]
    );
}

#[test]
fn yield_star_return_value_is_done_value() {
    assert_eq!(
        run_js(
            r#"
function* gen() {
    const result = yield* (function*() {
        yield 1; yield 2;
        return "final";
    })();
    yield result; // "final" from delegated generator's done value
}
console.log([...gen()].join(","));
"#
        ),
        vec!["1,2,final"]
    );
}

// ── state machine pattern ─────────────────────────────────────────────────────

#[test]
fn generator_as_state_machine() {
    assert_eq!(
        run_js(
            r#"
function* trafficLight() {
    while (true) {
        yield "red";
        yield "green";
        yield "yellow";
    }
}
const light = trafficLight();
const states = [];
for (let i = 0; i < 5; i++) states.push(light.next().value);
console.log(states.join(","));
"#
        ),
        vec!["red,green,yellow,red,green"]
    );
}

// ── lazy sequence ─────────────────────────────────────────────────────────────

#[test]
fn generator_lazy_map_filter() {
    assert_eq!(
        run_js(
            r#"
function* lazyMap(iter, fn) {
    for (const v of iter) yield fn(v);
}
function* lazyFilter(iter, pred) {
    for (const v of iter) if (pred(v)) yield v;
}
function* range(n) { for (let i = 0; i < n; i++) yield i; }
function take(iter, n) {
    const result = [];
    for (const v of iter) { result.push(v); if (result.length >= n) break; }
    return result;
}

const pipeline = lazyFilter(
    lazyMap(range(100), x => x * x),
    x => x % 2 === 0
);
console.log(take(pipeline, 5).join(","));
"#
        ),
        vec!["0,4,16,36,64"]
    );
}

// ── producer/consumer ─────────────────────────────────────────────────────────

#[test]
fn generator_producer_consumer() {
    assert_eq!(
        run_js(
            r#"
function* producer() {
    const items = [1, 2, 3, 4, 5];
    for (const item of items) {
        const doubled = yield item;
        if (doubled !== undefined) {
            // consumer sent back a value
        }
    }
}

const gen = producer();
const results = [];
let next = gen.next();
while (!next.done) {
    results.push(next.value * 2);
    next = gen.next();
}
console.log(results.join(","));
"#
        ),
        vec!["2,4,6,8,10"]
    );
}

// ── symbol.iterator on generator ──────────────────────────────────────────────

#[test]
fn generator_is_both_iterable_and_iterator() {
    assert_eq!(
        run_js(
            r#"
function* gen() { yield 1; yield 2; }
const g = gen();
// Generator has Symbol.iterator that returns itself
console.log(g[Symbol.iterator]() === g);
// So it can be used in for...of after partially consuming
g.next(); // consume 1
const remaining = [...g]; // consume rest
console.log(remaining.join(","));
"#
        ),
        vec!["true", "2"]
    );
}

// ── generator with try/catch/finally ──────────────────────────────────────────

#[test]
fn generator_finally_runs_on_early_return() {
    assert_eq!(
        run_js(
            r#"
const log = [];
function* gen() {
    try {
        yield 1;
        yield 2;
    } finally {
        log.push("finally");
    }
}
const g = gen();
g.next();
g.return("stop");
console.log(log.join(","));
"#
        ),
        vec!["finally"]
    );
}

// ── recursive generator ───────────────────────────────────────────────────────

#[test]
fn recursive_tree_traversal_via_generator() {
    assert_eq!(
        run_js(
            r#"
function* walk(node) {
    yield node.value;
    if (node.left) yield* walk(node.left);
    if (node.right) yield* walk(node.right);
}
const tree = {
    value: 1,
    left: { value: 2, left: { value: 4, left: null, right: null }, right: null },
    right: { value: 3, left: null, right: null }
};
console.log([...walk(tree)].join(","));
"#
        ),
        vec!["1,2,4,3"]
    );
}

#[test]
fn generator_throw_on_completed_generator_throws_reason() {
    assert_eq!(
        run_js(
            r#"
function* gen() { yield 1; }
const g = gen();
g.next();
g.next();
try {
    g.throw("custom_err");
} catch (e) {
    console.log(e);
}
"#
        ),
        vec!["custom_err"]
    );
}
