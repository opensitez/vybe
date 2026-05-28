/// Functional programming — monads, functors, point-free style

use super::helpers::run_js;

#[test]
fn maybe_monad() {
    assert_eq!(run_js(r#"
class Maybe {
    constructor(v) { this._v = v; }
    static of(v) { return new Maybe(v); }
    isNothing() { return this._v == null; }
    map(fn) { return this.isNothing() ? this : Maybe.of(fn(this._v)); }
    getOrElse(def) { return this.isNothing() ? def : this._v; }
}
const result1 = Maybe.of(5).map(x => x * 2).map(x => x + 1).getOrElse(0);
const result2 = Maybe.of(null).map(x => x * 2).getOrElse(-1);
console.log(result1);
console.log(result2);
"#), vec!["11", "-1"]);
}

#[test]
fn either_monad() {
    assert_eq!(run_js(r#"
class Right {
    constructor(v) { this._v = v; }
    map(fn) { return new Right(fn(this._v)); }
    fold(_, f) { return f(this._v); }
}
class Left {
    constructor(v) { this._v = v; }
    map(_) { return this; }
    fold(f, _) { return f(this._v); }
}
const safe = (f, onError) => v => {
    try { return new Right(f(v)); }
    catch (e) { return new Left(onError(e)); }
};
const parseJSON = safe(JSON.parse, e => e.message);
const good = parseJSON('{"x":1}').map(o => o.x).fold(e => 0, v => v);
const bad = parseJSON("not json").fold(e => -1, v => v);
console.log(good);
console.log(bad);
"#), vec!["1", "-1"]);
}

#[test]
fn reader_monad_simple() {
    assert_eq!(run_js(r#"
class Reader {
    constructor(fn) { this.run = fn; }
    map(fn) { return new Reader(env => fn(this.run(env))); }
    flatMap(fn) { return new Reader(env => fn(this.run(env)).run(env)); }
    static of(v) { return new Reader(_ => v); }
    static ask() { return new Reader(env => env); }
}
const greet = Reader.ask()
    .map(env => env.greeting)
    .map(g => g + " World");
console.log(greet.run({ greeting: "Hello" }));
console.log(greet.run({ greeting: "Hi" }));
"#), vec!["Hello World", "Hi World"]);
}

#[test]
fn state_monad_counter() {
    assert_eq!(run_js(r#"
class State {
    constructor(fn) { this.run = fn; }
    map(fn) { return new State(s => { const [v, ns] = this.run(s); return [fn(v), ns]; }); }
    flatMap(fn) { return new State(s => { const [v, ns] = this.run(s); return fn(v).run(ns); }); }
    static of(v) { return new State(s => [v, s]); }
    static get() { return new State(s => [s, s]); }
    static put(s) { return new State(_ => [null, s]); }
}
const increment = State.get().flatMap(n => State.put(n + 1).flatMap(() => State.get()));
const [value, finalState] = increment.flatMap(() => increment).run(0);
console.log(value);
console.log(finalState);
"#), vec!["2", "2"]);
}

#[test]
fn functor_laws() {
    assert_eq!(run_js(r#"
class Box {
    constructor(v) { this._v = v; }
    map(fn) { return new Box(fn(this._v)); }
    value() { return this._v; }
}
const identity = x => x;
const double = x => x * 2;
const addOne = x => x + 1;
// Identity law: map(id) === id
console.log(new Box(5).map(identity).value() === new Box(5).value());
// Composition law: map(f).map(g) === map(g(f(x)))
const a = new Box(5).map(double).map(addOne).value();
const b = new Box(5).map(x => addOne(double(x))).value();
console.log(a === b);
"#), vec!["true", "true"]);
}

#[test]
fn compose_and_pipe() {
    assert_eq!(run_js(r#"
const compose = (...fns) => x => fns.reduceRight((v, f) => f(v), x);
const pipe = (...fns) => x => fns.reduce((v, f) => f(v), x);
const double = x => x * 2;
const addTen = x => x + 10;
const square = x => x * x;
const composed = compose(square, addTen, double);  // double -> addTen -> square
const piped = pipe(double, addTen, square);          // same
console.log(composed(5));  // (5*2+10)^2 = 400
console.log(piped(5));     // same
"#), vec!["400", "400"]);
}

#[test]
fn partial_application() {
    assert_eq!(run_js(r#"
function partial(fn, ...args) {
    return (...rest) => fn(...args, ...rest);
}
const add = (a, b, c) => a + b + c;
const add10 = partial(add, 10);
const add10and20 = partial(add, 10, 20);
console.log(add10(5, 3));
console.log(add10and20(7));
"#), vec!["18", "37"]);
}

#[test]
fn transducer_pattern() {
    assert_eq!(run_js(r#"
const map = fn => reducer => (acc, val) => reducer(acc, fn(val));
const filter = pred => reducer => (acc, val) => pred(val) ? reducer(acc, val) : acc;
const append = (acc, val) => { acc.push(val); return acc; };

const xform = [
    filter(x => x % 2 === 0),
    map(x => x * x)
].reduce((a, b) => b(a), append);

const result = [1,2,3,4,5,6].reduce(xform, []);
console.log(result.join(","));
"#), vec!["4,16,36"]);
}

#[test]
fn lens_pattern() {
    assert_eq!(run_js(r#"
const lens = (getter, setter) => ({ get: getter, set: setter });
const view = (l, obj) => l.get(obj);
const set = (l, val, obj) => l.set(val, obj);
const over = (l, fn, obj) => set(l, fn(view(l, obj)), obj);

const nameLens = lens(o => o.name, (v, o) => ({...o, name: v}));
const person = { name: "Alice", age: 30 };
console.log(view(nameLens, person));
const updated = over(nameLens, n => n.toUpperCase(), person);
console.log(updated.name);
console.log(person.name);
"#), vec!["Alice", "ALICE", "Alice"]);
}

#[test]
fn continuation_passing_style() {
    assert_eq!(run_js(r#"
function addCPS(a, b, k) { k(a + b); }
function multiplyCPS(a, b, k) { k(a * b); }
function sqrtCPS(n, k) { k(Math.sqrt(n)); }

addCPS(3, 4, sum =>
    multiplyCPS(sum, 2, product =>
        sqrtCPS(product, result =>
            console.log(Math.round(result * 100) / 100)
        )
    )
);
"#), vec!["3.74"]);
}

#[test]
fn memoize_recursive() {
    assert_eq!(run_js(r#"
function memoize(fn) {
    const cache = new Map();
    return (...args) => {
        const key = JSON.stringify(args);
        if (!cache.has(key)) cache.set(key, fn(...args));
        return cache.get(key);
    };
}
let calls = 0;
const square = memoize(n => { calls++; return n * n; });
console.log(square(5));
console.log(square(5));
console.log(square(3));
const c = calls;
square(5);
square(3);
console.log(calls === c);
"#), vec!["25", "25", "9", "true"]);
}
