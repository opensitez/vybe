use super::helpers::run_js;

// ── Generator delegation (yield*) ────────────────────────
#[test]
fn yield_star_delegates_to_array() {
    assert_eq!(run_js(r#"
function* gen() { yield* [1, 2, 3]; }
console.log([...gen()].join(","));
"#), vec!["1,2,3"]);
}

#[test]
fn yield_star_delegates_to_another_generator() {
    assert_eq!(run_js(r#"
function* inner() { yield "a"; yield "b"; }
function* outer() { yield 1; yield* inner(); yield 2; }
console.log([...outer()].join(","));
"#), vec!["1,a,b,2"]);
}

#[test]
fn yield_star_return_value() {
    assert_eq!(run_js(r#"
function* inner() { yield 1; return "done"; }
function* outer() {
  const result = yield* inner();
  yield result;
}
console.log([...outer()].join(","));
"#), vec!["1,done"]);
}

#[test]
fn yield_star_string_iterates_chars() {
    assert_eq!(run_js(r#"
function* charGen() { yield* "abc"; }
console.log([...charGen()].join(","));
"#), vec!["a,b,c"]);
}

// ── Generator throw/return ────────────────────────────────
#[test]
fn generator_return_method_stops() {
    assert_eq!(run_js(r#"
function* gen() { yield 1; yield 2; yield 3; }
const g = gen();
console.log(g.next().value);
console.log(g.return(99).value);
console.log(g.next().done);
"#), vec!["1", "99", "true"]);
}

#[test]
fn generator_throw_caught_inside() {
    assert_eq!(run_js(r#"
function* gen() {
  try {
    yield 1;
    yield 2;
  } catch (e) {
    yield "caught: " + e;
  }
}
const g = gen();
console.log(g.next().value);
console.log(g.throw("oops").value);
console.log(g.next().done);
"#), vec!["1", "caught: oops", "true"]);
}

#[test]
fn generator_finally_on_return() {
    assert_eq!(run_js(r#"
const steps = [];
function* gen() {
  try {
    yield 1;
    yield 2;
  } finally {
    steps.push("finally");
  }
}
const g = gen();
g.next();
g.return("end");
console.log(steps.join(","));
"#), vec!["finally"]);
}

// ── Generator as iterator ─────────────────────────────────
#[test]
fn generator_implements_iterator_protocol() {
    assert_eq!(run_js(r#"
function* counter() { let n = 0; while (true) yield n++; }
const it = counter();
console.log(it.next().value);
console.log(it.next().value);
console.log(it.next().value);
"#), vec!["0", "1", "2"]);
}

#[test]
fn generator_is_both_iterable_and_iterator() {
    assert_eq!(run_js(r#"
function* gen() { yield 1; yield 2; }
const g = gen();
console.log(g[Symbol.iterator]() === g);
"#), vec!["true"]);
}

#[test]
fn generator_for_of_loop() {
    assert_eq!(run_js(r#"
function* evens(limit) {
  for (let i = 2; i <= limit; i += 2) yield i;
}
const result = [];
for (const v of evens(10)) result.push(v);
console.log(result.join(","));
"#), vec!["2,4,6,8,10"]);
}

// ── send values into generators ───────────────────────────
#[test]
fn generator_receives_sent_values() {
    assert_eq!(run_js(r#"
function* adder() {
  let sum = 0;
  while (true) {
    const n = yield sum;
    if (n === null) break;
    sum += n;
  }
  return sum;
}
const g = adder();
g.next();
g.next(5);
g.next(3);
const { value } = g.next(null);
console.log(value);
"#), vec!["8"]);
}

#[test]
fn generator_first_next_arg_ignored() {
    assert_eq!(run_js(r#"
function* gen() {
  const x = yield "first";
  yield x * 2;
}
const g = gen();
g.next("ignored");
const { value } = g.next(10);
console.log(value);
"#), vec!["20"]);
}

// ── Infinite generators ───────────────────────────────────
#[test]
fn generator_fibonacci_sequence() {
    assert_eq!(run_js(r#"
function* fib() {
  let [a, b] = [0, 1];
  while (true) { yield a; [a, b] = [b, a + b]; }
}
const g = fib();
const first8 = [];
for (let i = 0; i < 8; i++) first8.push(g.next().value);
console.log(first8.join(","));
"#), vec!["0,1,1,2,3,5,8,13"]);
}

#[test]
fn generator_take_helper() {
    assert_eq!(run_js(r#"
function* naturals() { let n = 1; while (true) yield n++; }
function take(gen, n) {
  const result = [];
  for (const v of gen) { result.push(v); if (result.length === n) break; }
  return result;
}
console.log(take(naturals(), 5).join(","));
"#), vec!["1,2,3,4,5"]);
}

// ── Generator return value ────────────────────────────────
#[test]
fn generator_return_statement_value() {
    assert_eq!(run_js(r#"
function* gen() { yield 1; yield 2; return "final"; }
const g = gen();
g.next(); g.next();
const { value, done } = g.next();
console.log(value, done);
"#), vec!["final true"]);
}

#[test]
fn spread_ignores_generator_return_value() {
    assert_eq!(run_js(r#"
function* gen() { yield 1; yield 2; return 99; }
const arr = [...gen()];
console.log(arr.join(","));
"#), vec!["1,2"]);
}

// ── Async generators ──────────────────────────────────────
#[test]
fn async_generator_basic() {
    assert_eq!(run_js(r#"
async function* asyncRange(start, end) {
  for (let i = start; i <= end; i++) yield i;
}
async function collect() {
  const result = [];
  for await (const v of asyncRange(1, 4)) result.push(v);
  console.log(result.join(","));
}
collect();
"#), vec!["1,2,3,4"]);
}

#[test]
fn async_generator_yield_promises() {
    assert_eq!(run_js(r#"
async function* gen() {
  yield await Promise.resolve(1);
  yield await Promise.resolve(2);
}
async function run() {
  const result = [];
  for await (const v of gen()) result.push(v);
  console.log(result.join(","));
}
run();
"#), vec!["1,2"]);
}

// ── Generator composition ─────────────────────────────────
#[test]
fn generators_compose_pipeline() {
    assert_eq!(run_js(r#"
function* map(iter, fn) { for (const v of iter) yield fn(v); }
function* filter(iter, pred) { for (const v of iter) if (pred(v)) yield v; }
function* range(n) { for (let i = 1; i <= n; i++) yield i; }

const pipeline = filter(map(range(10), x => x * x), x => x % 2 === 0);
const result = [...pipeline];
console.log(result.join(","));
"#), vec!["4,16,36,64,100"]);
}

#[test]
fn generator_chained_delegation() {
    assert_eq!(run_js(r#"
function* a() { yield 1; yield 2; }
function* b() { yield* a(); yield 3; }
function* c() { yield* b(); yield 4; }
console.log([...c()].join(","));
"#), vec!["1,2,3,4"]);
}

#[test]
fn generator_early_break_in_for_of() {
    assert_eq!(run_js(r#"
const visited = [];
function* gen() {
  try {
    yield 1; yield 2; yield 3;
  } finally {
    visited.push("cleanup");
  }
}
for (const v of gen()) {
  visited.push(v);
  if (v === 2) break;
}
console.log(visited.join(","));
"#), vec!["1,2,cleanup"]);
}

#[test]
fn generator_symbol_iterator_makes_reusable() {
    assert_eq!(run_js(r#"
const iterable = {
  [Symbol.iterator]: function* () { yield "x"; yield "y"; yield "z"; }
};
const r1 = [...iterable].join(",");
const r2 = [...iterable].join(",");
console.log(r1);
console.log(r1 === r2);
"#), vec!["x,y,z", "true"]);
}

#[test]
fn generator_with_object_destructuring() {
    assert_eq!(run_js(r#"
function* pairs() {
  yield { key: "a", val: 1 };
  yield { key: "b", val: 2 };
}
const result = [];
for (const { key, val } of pairs()) result.push(key + "=" + val);
console.log(result.join(","));
"#), vec!["a=1,b=2"]);
}
