/// Generator advanced — infinite sequences, pipelines, coroutine, tree traversal
use super::helpers::run_js;

#[test]
fn infinite_sequence_take() {
    assert_eq!(
        run_js(
            r#"
function* naturals() {
    let n = 1;
    while (true) yield n++;
}
function take(n, gen) {
    const result = [];
    for (const v of gen) {
        result.push(v);
        if (result.length === n) break;
    }
    return result;
}
console.log(take(5, naturals()).join(","));
"#
        ),
        vec!["1,2,3,4,5"]
    );
}

#[test]
fn fibonacci_generator() {
    assert_eq!(
        run_js(
            r#"
function* fib() {
    let a = 0, b = 1;
    while (true) {
        yield a;
        [a, b] = [b, a + b];
    }
}
const gen = fib();
const first8 = [];
for (let i = 0; i < 8; i++) first8.push(gen.next().value);
console.log(first8.join(","));
"#
        ),
        vec!["0,1,1,2,3,5,8,13"]
    );
}

#[test]
fn generator_as_lazy_filter() {
    assert_eq!(
        run_js(
            r#"
function* filter(pred, gen) {
    for (const v of gen) if (pred(v)) yield v;
}
function* range(n) { for (let i = 0; i < n; i++) yield i; }
const evens = [...filter(x => x % 2 === 0, range(10))];
console.log(evens.join(","));
"#
        ),
        vec!["0,2,4,6,8"]
    );
}

#[test]
fn generator_as_lazy_map() {
    assert_eq!(
        run_js(
            r#"
function* map(fn, gen) {
    for (const v of gen) yield fn(v);
}
function* range(n) { for (let i = 1; i <= n; i++) yield i; }
const squares = [...map(x => x * x, range(5))];
console.log(squares.join(","));
"#
        ),
        vec!["1,4,9,16,25"]
    );
}

#[test]
fn generator_tree_traversal_preorder() {
    assert_eq!(
        run_js(
            r#"
function* preorder(node) {
    if (!node) return;
    yield node.val;
    yield* preorder(node.left);
    yield* preorder(node.right);
}
const tree = {
    val: 1,
    left: { val: 2, left: { val: 4, left: null, right: null }, right: null },
    right: { val: 3, left: null, right: null }
};
console.log([...preorder(tree)].join(","));
"#
        ),
        vec!["1,2,4,3"]
    );
}

#[test]
fn generator_coroutine_via_send() {
    assert_eq!(
        run_js(
            r#"
function* accumulator() {
    let sum = 0;
    while (true) {
        const n = yield sum;
        if (n === null) break;
        sum += n;
    }
}
const gen = accumulator();
gen.next();       // start
gen.next(10);
gen.next(20);
const result = gen.next(5);
console.log(result.value); // 35
"#
        ),
        vec!["35"]
    );
}

#[test]
fn generator_yield_star_with_return_value() {
    assert_eq!(
        run_js(
            r#"
function* inner() {
    yield 1;
    yield 2;
    return "inner done";
}
function* outer() {
    const result = yield* inner();
    console.log(result); // return value of inner
    yield 3;
}
console.log([...outer()].join(","));
"#
        ),
        vec!["inner done", "1,2,3"]
    );
}

#[test]
fn generator_pipeline_flattening() {
    assert_eq!(
        run_js(
            r#"
function* flatten(arr, depth = 1) {
    for (const item of arr) {
        if (Array.isArray(item) && depth > 0) yield* flatten(item, depth - 1);
        else yield item;
    }
}
const nested = [1, [2, [3, [4]]], 5];
// JSON.stringify (not join): join would render the still-nested [4] as
// "4" (Array.prototype.join calls toString on elements), hiding the very
// thing this test checks — that depth-limited flatten leaves [4] nested.
console.log(JSON.stringify([...flatten(nested, 2)]));
"#
        ),
        vec!["[1,2,3,[4],5]"]
    );
}

#[test]
fn generator_zip_two_iterables() {
    assert_eq!(
        run_js(
            r#"
function* zip(a, b) {
    const itA = a[Symbol.iterator]();
    const itB = b[Symbol.iterator]();
    while (true) {
        const rA = itA.next();
        const rB = itB.next();
        if (rA.done || rB.done) break;
        yield [rA.value, rB.value];
    }
}
const zipped = [...zip([1, 2, 3], ["a", "b", "c"])];
console.log(zipped.map(([a, b]) => a + b).join(","));
"#
        ),
        vec!["1a,2b,3c"]
    );
}

#[test]
fn generator_return_executes_finally_block() {
    assert_eq!(
        run_js(
            r#"
function* gen() {
    try {
        yield 1;
        yield 2;
    } finally {
        console.log("finally");
    }
}
const g = gen();
g.next();
const r = g.return("done");
console.log(r.value + "|" + r.done);
"#
        ),
        vec!["finally", "done|true"]
    );
}
