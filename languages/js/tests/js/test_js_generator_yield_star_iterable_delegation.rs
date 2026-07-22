use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Generator `yield*` Iterable Delegation Protocol
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_generator_yield_star_delegation_to_array() {
    let src = r#"
function* gen() {
    yield* [10, 20, 30];
}
console.log([...gen()].join(","));
"#;
    assert_eq!(run_js(src), vec!["10,20,30"]);
}

#[test]
fn test_js_generator_yield_star_delegation_to_subgenerator() {
    let src = r#"
function* inner() {
    yield "a";
    yield "b";
    return "innerDone";
}
function* outer() {
    const res = yield* inner();
    yield res;
}
console.log([...outer()].join(","));
"#;
    assert_eq!(run_js(src), vec!["a,b,innerDone"]); // Return value of inner generator is received by yield* expression!
}

#[test]
fn test_js_generator_yield_star_passing_arguments_into_next() {
    let src = r#"
function* inner() {
    const x = yield 1;
    const y = yield x * 2;
    return y * 3;
}
function* outer() {
    const ret = yield* inner();
    yield "outerRet:" + ret;
}
const g = outer();
console.log(g.next().value); // 1
console.log(g.next(10).value); // 20
console.log(g.next(5).value); // "outerRet:15"
"#;
    assert_eq!(run_js(src), vec!["1", "20", "outerRet:15"]);
}

#[test]
fn test_js_generator_yield_star_string_delegation() {
    let src = r#"
function* gen() {
    yield* "JS";
}
console.log([...gen()].join("-"));
"#;
    assert_eq!(run_js(src), vec!["J-S"]);
}

#[test]
fn test_js_generator_yield_star_set_and_map_delegation() {
    let src = r#"
function* gen() {
    yield* new Set(["X", "Y"]);
}
console.log([...gen()].join(","));
"#;
    assert_eq!(run_js(src), vec!["X,Y"]);
}

#[test]
fn test_js_generator_yield_star_custom_iterable() {
    let src = r#"
const customIterable = {
    *[Symbol.iterator]() { yield 100; yield 200; }
};
function* gen() {
    yield* customIterable;
}
console.log([...gen()].join(","));
"#;
    assert_eq!(run_js(src), vec!["100,200"]);
}

#[test]
fn test_js_generator_yield_star_return_propagation() {
    let src = r#"
let innerReturned = false;
function* inner() {
    try {
        yield 1;
        yield 2;
    } finally {
        innerReturned = true;
    }
}
function* outer() {
    yield* inner();
}
const g = outer();
g.next();
g.return("EarlyReturn");
console.log(innerReturned);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_generator_yield_star_throw_propagation_into_inner() {
    let src = r#"
function* inner() {
    try {
        yield "innerVal";
    } catch (e) {
        yield "innerCaught: " + e.message;
    }
}
function* outer() {
    yield* inner();
}
const g = outer();
g.next();
console.log(g.throw(new Error("DelegatedError")).value);
"#;
    assert_eq!(run_js(src), vec!["innerCaught: DelegatedError"]);
}

#[test]
fn test_js_generator_yield_star_non_iterable_throws_typeerror() {
    let src = r#"
function* gen() {
    yield* 12345;
}
try {
    [...gen()];
} catch (e) {
    console.log("yield* Non-Iterable TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["yield* Non-Iterable TypeError"]);
}

#[test]
fn test_js_generator_yield_star_null_or_undefined_throws_typeerror() {
    let src = r#"
function* gen() {
    yield* null;
}
try {
    [...gen()];
} catch (e) {
    console.log("yield* Null TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["yield* Null TypeError"]);
}

#[test]
fn test_js_generator_yield_star_chained_delegation() {
    let src = r#"
function* gen1() { yield 1; }
function* gen2() { yield* gen1(); yield 2; }
function* gen3() { yield* gen2(); yield 3; }
console.log([...gen3()].join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2,3"]);
}

#[test]
fn test_js_generator_yield_star_arguments_object_delegation() {
    let src = r#"
function* gen() {
    yield* arguments;
}
function test(a, b) {
    return [...gen(a, b)];
}
console.log(test("first", "second").join(","));
"#;
    assert_eq!(run_js(src), vec!["first,second"]);
}

#[test]
fn test_js_generator_yield_star_typed_array_delegation() {
    let src = r#"
function* gen() {
    yield* new Uint8Array([5, 10, 15]);
}
console.log([...gen()].join(","));
"#;
    assert_eq!(run_js(src), vec!["5,10,15"]);
}

#[test]
fn test_js_generator_yield_star_iterator_without_throw_method() {
    let src = r#"
const customIter = {
    [Symbol.iterator]() {
        return {
            next() { return { value: "x", done: false }; }
        };
    }
};
function* gen() {
    yield* customIter;
}
const g = gen();
g.next();
try {
    g.throw(new Error("NoThrowMethod"));
} catch (e) {
    console.log(e.message);
}
"#;
    assert_eq!(run_js(src), vec!["NoThrowMethod"]);
}

#[test]
fn test_js_generator_yield_star_iterator_without_return_method() {
    let src = r#"
const customIter = {
    [Symbol.iterator]() {
        return {
            next() { return { value: "y", done: false }; }
        };
    }
};
function* gen() {
    yield* customIter;
}
const g = gen();
g.next();
const ret = g.return("NoReturnMethod");
console.log(ret.value);
"#;
    assert_eq!(run_js(src), vec!["NoReturnMethod"]);
}

#[test]
fn test_js_generator_yield_star_with_tree_traversal() {
    let src = r#"
const tree = {
    val: 1,
    children: [
        { val: 2, children: [] },
        { val: 3, children: [{ val: 4, children: [] }] }
    ]
};
function* traverse(node) {
    yield node.val;
    for (const child of node.children) {
        yield* traverse(child);
    }
}
console.log([...traverse(tree)].join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2,3,4"]);
}

#[test]
fn test_js_generator_yield_star_result_expression_in_operator() {
    let src = r#"
function* inner() { return 100; }
function* outer() {
    const val = (yield* inner()) * 2;
    yield val;
}
console.log([...outer()][0]);
"#;
    assert_eq!(run_js(src), vec!["200"]);
}

#[test]
fn test_js_generator_yield_star_empty_generator_returns_value() {
    let src = r#"
function* emptyGen() { return "EmptyVal"; }
function* outer() {
    const res = yield* emptyGen();
    yield res;
}
console.log([...outer()].join(","));
"#;
    assert_eq!(run_js(src), vec!["EmptyVal"]);
}

#[test]
fn test_js_generator_yield_star_delegation_to_map_entries() {
    let src = r#"
const map = new Map([["k1", "v1"], ["k2", "v2"]]);
function* gen() {
    yield* map.entries();
}
console.log([...gen()].map(pair => pair.join("=")).join(","));
"#;
    assert_eq!(run_js(src), vec!["k1=v1,k2=v2"]);
}

#[test]
fn test_js_generator_yield_star_delegation_to_set_values() {
    let src = r#"
const set = new Set(["a", "b"]);
function* gen() {
    yield* set.values();
}
console.log([...gen()].join(","));
"#;
    assert_eq!(run_js(src), vec!["a,b"]);
}
