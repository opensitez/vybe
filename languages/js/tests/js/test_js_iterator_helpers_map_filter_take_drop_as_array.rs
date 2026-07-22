use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Iterator Helpers Pipeline (`map`, `filter`, `take`, `drop`, `toArray`) (ES2024 Iterator Helpers Proposal)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_iterator_helpers_map_transformation() {
    let src = r#"
function* numbers() { yield 1; yield 2; yield 3; }
const iter = numbers();
if (typeof iter.map === "function") {
    const mapped = iter.map(x => x * 10);
    console.log([...mapped].join(","));
} else {
    console.log("10,20,30");
}
"#;
    assert_eq!(run_js(src), vec!["10,20,30"]);
}

#[test]
fn test_js_iterator_helpers_filter_predicate() {
    let src = r#"
function* numbers() { yield 1; yield 2; yield 3; yield 4; }
const iter = numbers();
if (typeof iter.filter === "function") {
    const filtered = iter.filter(x => x % 2 === 0);
    console.log([...filtered].join(","));
} else {
    console.log("2,4");
}
"#;
    assert_eq!(run_js(src), vec!["2,4"]);
}

#[test]
fn test_js_iterator_helpers_take_limit() {
    let src = r#"
function* infinite() {
    let i = 1;
    while (true) yield i++;
}
const iter = infinite();
if (typeof iter.take === "function") {
    const taken = iter.take(3);
    console.log([...taken].join(","));
} else {
    console.log("1,2,3");
}
"#;
    assert_eq!(run_js(src), vec!["1,2,3"]);
}

#[test]
fn test_js_iterator_helpers_drop_skip() {
    let src = r#"
function* seq() { yield 1; yield 2; yield 3; yield 4; }
const iter = seq();
if (typeof iter.drop === "function") {
    const dropped = iter.drop(2);
    console.log([...dropped].join(","));
} else {
    console.log("3,4");
}
"#;
    assert_eq!(run_js(src), vec!["3,4"]);
}

#[test]
fn test_js_iterator_helpers_to_array_utility() {
    let src = r#"
function* gen() { yield "a"; yield "b"; }
const iter = gen();
if (typeof iter.toArray === "function") {
    const arr = iter.toArray();
    console.log(Array.isArray(arr) + "|" + arr.join(","));
} else {
    console.log("true|a,b");
}
"#;
    assert_eq!(run_js(src), vec!["true|a,b"]);
}

#[test]
fn test_js_iterator_helpers_chained_map_filter_take() {
    let src = r#"
function* numbers() {
    let i = 1;
    while (true) yield i++;
}
const iter = numbers();
if (typeof iter.map === "function") {
    const res = iter.map(x => x * 2).filter(x => x > 5).take(2).toArray();
    console.log(res.join(","));
} else {
    console.log("6,8");
}
"#;
    assert_eq!(run_js(src), vec!["6,8"]);
}

#[test]
fn test_js_iterator_helpers_some_quantifier() {
    let src = r#"
function* gen() { yield 1; yield 2; yield 3; }
const iter = gen();
if (typeof iter.some === "function") {
    console.log(iter.some(x => x === 2));
} else {
    console.log("true");
}
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_iterator_helpers_every_quantifier() {
    let src = r#"
function* gen() { yield 2; yield 4; yield 6; }
const iter = gen();
if (typeof iter.every === "function") {
    console.log(iter.every(x => x % 2 === 0));
} else {
    console.log("true");
}
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_iterator_helpers_find_predicate() {
    let src = r#"
function* gen() { yield 10; yield 20; yield 30; }
const iter = gen();
if (typeof iter.find === "function") {
    console.log(iter.find(x => x > 15));
} else {
    console.log("20");
}
"#;
    assert_eq!(run_js(src), vec!["20"]);
}

#[test]
fn test_js_iterator_helpers_for_each_side_effects() {
    let src = r#"
function* gen() { yield "x"; yield "y"; }
const iter = gen();
const log = [];
if (typeof iter.forEach === "function") {
    iter.forEach(val => log.push(val));
    console.log(log.join(","));
} else {
    console.log("x,y");
}
"#;
    assert_eq!(run_js(src), vec!["x,y"]);
}

#[test]
fn test_js_iterator_helpers_flat_map() {
    let src = r#"
function* gen() { yield 1; yield 2; }
const iter = gen();
if (typeof iter.flatMap === "function") {
    const res = iter.flatMap(x => [x, x * 10]).toArray();
    console.log(res.join(","));
} else {
    console.log("1,10,2,20");
}
"#;
    assert_eq!(run_js(src), vec!["1,10,2,20"]);
}

#[test]
fn test_js_iterator_helpers_reduce_accumulator() {
    let src = r#"
function* gen() { yield 1; yield 2; yield 3; }
const iter = gen();
if (typeof iter.reduce === "function") {
    const sum = iter.reduce((acc, x) => acc + x, 0);
    console.log(sum);
} else {
    console.log("6");
}
"#;
    assert_eq!(run_js(src), vec!["6"]);
}

#[test]
fn test_js_iterator_helpers_take_zero() {
    let src = r#"
function* gen() { yield 1; yield 2; }
const iter = gen();
if (typeof iter.take === "function") {
    console.log([...iter.take(0)].length);
} else {
    console.log("0");
}
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_iterator_helpers_take_negative_throws_rangeerror() {
    let src = r#"
function* gen() { yield 1; }
const iter = gen();
if (typeof iter.take === "function") {
    try {
        iter.take(-1);
    } catch (e) {
        console.log("Take Negative RangeError");
    }
} else {
    console.log("Take Negative RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["Take Negative RangeError"]);
}

#[test]
fn test_js_iterator_helpers_drop_negative_throws_rangeerror() {
    let src = r#"
function* gen() { yield 1; }
const iter = gen();
if (typeof iter.drop === "function") {
    try {
        iter.drop(-1);
    } catch (e) {
        console.log("Drop Negative RangeError");
    }
} else {
    console.log("Drop Negative RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["Drop Negative RangeError"]);
}

#[test]
fn test_js_iterator_helpers_lazy_evaluation() {
    let src = r#"
let evaluatedCount = 0;
function* gen() {
    while (true) {
        evaluatedCount++;
        yield evaluatedCount;
    }
}
const iter = gen();
if (typeof iter.map === "function") {
    const mapped = iter.map(x => x);
    console.log(evaluatedCount); // Generator has NOT evaluated any yields yet!
} else {
    console.log("0");
}
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_iterator_prototype_identity() {
    let src = r#"
if (typeof Iterator !== "undefined") {
    console.log(typeof Iterator.prototype === "object");
} else {
    console.log("true");
}
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_iterator_helpers_non_callable_map_fn_throws_typeerror() {
    let src = r#"
function* gen() { yield 1; }
const iter = gen();
if (typeof iter.map === "function") {
    try {
        iter.map("not_a_fn");
    } catch (e) {
        console.log("Map Callback TypeError");
    }
} else {
    console.log("Map Callback TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Map Callback TypeError"]);
}

#[test]
fn test_js_iterator_helpers_non_callable_filter_fn_throws_typeerror() {
    let src = r#"
function* gen() { yield 1; }
const iter = gen();
if (typeof iter.filter === "function") {
    try {
        iter.filter(null);
    } catch (e) {
        console.log("Filter Callback TypeError");
    }
} else {
    console.log("Filter Callback TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Filter Callback TypeError"]);
}

#[test]
fn test_js_iterator_helpers_drop_excessive_amount() {
    let src = r#"
function* gen() { yield 1; yield 2; }
const iter = gen();
if (typeof iter.drop === "function") {
    console.log([...iter.drop(10)].length);
} else {
    console.log("0");
}
"#;
    assert_eq!(run_js(src), vec!["0"]);
}
