/// Generator delegation and advanced generator composition
use super::helpers::run_js;

#[test]
fn yield_star_delegates_to_generator() {
    assert_eq!(
        run_js(
            r#"
function* inner() { yield 1; yield 2; yield 3; }
function* outer() { yield 0; yield* inner(); yield 4; }
console.log([...outer()].join(","));
"#
        ),
        vec!["0,1,2,3,4"]
    );
}

#[test]
fn yield_star_return_value() {
    assert_eq!(
        run_js(
            r#"
function* sub() {
    yield 1;
    yield 2;
    return "done";
}
function* main() {
    const result = yield* sub();
    yield result;
}
console.log([...main()].join(","));
"#
        ),
        vec!["1,2,done"]
    );
}

#[test]
fn generator_pipeline() {
    assert_eq!(
        run_js(
            r#"
function* naturals(n = Infinity) {
    let i = 0;
    while (i < n) yield i++;
}
function* map(gen, fn) { for (const v of gen) yield fn(v); }
function* filter(gen, pred) { for (const v of gen) if (pred(v)) yield v; }
function* take(gen, n) { let i = 0; for (const v of gen) { if (i++ >= n) break; yield v; } }

const result = [...take(filter(map(naturals(), x => x*x), x => x % 2 === 0), 5)];
console.log(result.join(","));
"#
        ),
        vec!["0,4,16,36,64"]
    );
}

#[test]
fn generator_tree_dfs() {
    assert_eq!(
        run_js(
            r#"
function* dfs(node) {
    yield node.value;
    if (node.left) yield* dfs(node.left);
    if (node.right) yield* dfs(node.right);
}
const tree = {
    value: 1,
    left: { value: 2, left: { value: 4, left: null, right: null }, right: null },
    right: { value: 3, left: null, right: { value: 5, left: null, right: null } }
};
console.log([...dfs(tree)].join(","));
"#
        ),
        vec!["1,2,4,3,5"]
    );
}

#[test]
fn generator_as_state_machine() {
    assert_eq!(
        run_js(
            r#"
function* trafficLight() {
    while (true) {
        yield "green";
        yield "yellow";
        yield "red";
    }
}
const light = trafficLight();
const states = Array.from({length: 7}, () => light.next().value);
console.log(states.join(","));
"#
        ),
        vec!["green,yellow,red,green,yellow,red,green"]
    );
}

#[test]
fn recursive_generator_flatten() {
    assert_eq!(
        run_js(
            r#"
function* flatten(arr) {
    for (const item of arr) {
        if (Array.isArray(item)) yield* flatten(item);
        else yield item;
    }
}
const nested = [1, [2, [3, [4, [5]]]], 6];
console.log([...flatten(nested)].join(","));
"#
        ),
        vec!["1,2,3,4,5,6"]
    );
}

#[test]
fn generator_zip_two() {
    assert_eq!(
        run_js(
            r#"
function* zip(...iters) {
    const gen = iters.map(i => i[Symbol.iterator]());
    while (true) {
        const results = gen.map(g => g.next());
        if (results.some(r => r.done)) break;
        yield results.map(r => r.value);
    }
}
const pairs = [...zip([1,2,3], ["a","b","c"])];
console.log(pairs.map(p => p.join(":")).join(","));
"#
        ),
        vec!["1:a,2:b,3:c"]
    );
}

#[test]
fn generator_chunk() {
    assert_eq!(
        run_js(
            r#"
function* chunk(iter, size) {
    let batch = [];
    for (const item of iter) {
        batch.push(item);
        if (batch.length === size) { yield batch; batch = []; }
    }
    if (batch.length) yield batch;
}
const chunks = [...chunk([1,2,3,4,5,6,7], 3)];
console.log(chunks.length);
console.log(chunks[0].join(","));
console.log(chunks[2].join(","));
"#
        ),
        vec!["3", "1,2,3", "7"]
    );
}

#[test]
fn generator_memoize_via_cache() {
    assert_eq!(
        run_js(
            r#"
function* uniqueValues(gen) {
    const seen = new Set();
    for (const v of gen) {
        if (!seen.has(v)) { seen.add(v); yield v; }
    }
}
const input = [1, 2, 2, 3, 1, 4, 3, 5];
console.log([...uniqueValues(input)].join(","));
"#
        ),
        vec!["1,2,3,4,5"]
    );
}

#[test]
fn async_generator_paginate() {
    assert_eq!(
        run_js(
            r#"
async function* paginate(data, pageSize) {
    for (let i = 0; i < data.length; i += pageSize) {
        yield data.slice(i, i + pageSize);
    }
}
async function main() {
    const pages = [];
    for await (const page of paginate([1,2,3,4,5,6,7], 3)) {
        pages.push(page.join(","));
    }
    console.log(pages.length);
    console.log(pages[0]);
    console.log(pages[2]);
}
main();
"#
        ),
        vec!["3", "1,2,3", "7"]
    );
}
