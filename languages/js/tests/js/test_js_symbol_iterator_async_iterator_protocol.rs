use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `Symbol.iterator` & `Symbol.asyncIterator` Protocol Implementation
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_custom_symbol_iterator_for_of() {
    let src = r#"
const customCollection = {
    items: [10, 20, 30],
    [Symbol.iterator]() {
        let idx = 0;
        const items = this.items;
        return {
            next() {
                if (idx < items.length) {
                    return { value: items[idx++], done: false };
                }
                return { value: undefined, done: true };
            }
        };
    }
};
const res = [];
for (const val of customCollection) res.push(val);
console.log(res.join(","));
"#;
    assert_eq!(run_js(src), vec!["10,20,30"]);
}

#[test]
fn test_js_custom_symbol_async_iterator_for_await_of() {
    let src = r#"
const asyncSeq = {
    [Symbol.asyncIterator]() {
        let i = 1;
        return {
            async next() {
                if (i <= 3) {
                    return { value: await Promise.resolve(i++ * 10), done: false };
                }
                return { done: true };
            }
        };
    }
};
(async () => {
    const res = [];
    for await (const val of asyncSeq) res.push(val);
    console.log(res.join(","));
})();
"#;
    assert_eq!(run_js(src), vec!["10,20,30"]);
}

#[test]
fn test_js_symbol_iterator_array_spread() {
    let src = r#"
const range = {
    from: 1, to: 3,
    *[Symbol.iterator]() {
        for (let i = this.from; i <= this.to; i++) yield i;
    }
};
console.log([...range].join("-"));
"#;
    assert_eq!(run_js(src), vec!["1-2-3"]);
}

#[test]
fn test_js_symbol_iterator_destructuring() {
    let src = r#"
const tupleObj = {
    [Symbol.iterator]: function*() {
        yield "first";
        yield "second";
    }
};
const [a, b] = tupleObj;
console.log(`${a}:${b}`);
"#;
    assert_eq!(run_js(src), vec!["first:second"]);
}

#[test]
fn test_js_symbol_iterator_return_method_early_break() {
    let src = r#"
let cleanedUp = false;
const iterObj = {
    [Symbol.iterator]() {
        return {
            next() { return { value: 1, done: false }; },
            return() {
                cleanedUp = true;
                return { done: true };
            }
        };
    }
};
for (const val of iterObj) {
    if (val === 1) break; // Early break triggers return() method on iterator!
}
console.log(cleanedUp);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_symbol_iterator_throw_method_handling() {
    let src = r#"
function* gen() {
    try {
        yield 1;
    } catch (e) {
        yield "CaughtInGen: " + e.message;
    }
}
const g = gen();
g.next();
console.log(g.throw(new Error("ExternalError")).value);
"#;
    assert_eq!(run_js(src), vec!["CaughtInGen: ExternalError"]);
}

#[test]
fn test_js_symbol_async_iterator_fallback_from_sync_for_await() {
    let src = r#"
const syncIterable = [100, 200];
(async () => {
    const res = [];
    for await (const val of syncIterable) { // for-await-of falls back to Symbol.iterator wrapped in Promise if asyncIterator is missing
        res.push(val);
    }
    console.log(res.join(","));
})();
"#;
    assert_eq!(run_js(src), vec!["100,200"]);
}

#[test]
fn test_js_symbol_iterator_non_object_return_from_next_throws_typeerror() {
    let src = r#"
const badIter = {
    [Symbol.iterator]() {
        return { next() { return "not_an_object"; } };
    }
};
try {
    for (const _ of badIter);
} catch (e) {
    console.log("Iterator Next Non-Object TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Iterator Next Non-Object TypeError"]);
}

#[test]
fn test_js_symbol_iterator_not_a_function_throws_typeerror() {
    let src = r#"
const badIter = { [Symbol.iterator]: "not_a_function" };
try {
    for (const _ of badIter);
} catch (e) {
    console.log("Symbol.iterator Not Callable TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Symbol.iterator Not Callable TypeError"]);
}

#[test]
fn test_js_symbol_iterator_generator_method_concise_syntax() {
    let src = r#"
class Queue {
    #items = ["Q1", "Q2"];
    *[Symbol.iterator]() {
        yield* this.#items;
    }
}
console.log([...new Queue()].join(","));
"#;
    assert_eq!(run_js(src), vec!["Q1,Q2"]);
}

#[test]
fn test_js_symbol_async_iterator_async_generator_concise_syntax() {
    let src = r#"
class AsyncQueue {
    async *[Symbol.asyncIterator]() {
        yield await Promise.resolve("AQ1");
        yield await Promise.resolve("AQ2");
    }
}
(async () => {
    const items = [];
    for await (const x of new AsyncQueue()) items.push(x);
    console.log(items.join(","));
})();
"#;
    assert_eq!(run_js(src), vec!["AQ1,AQ2"]);
}

#[test]
fn test_js_symbol_iterator_prototype_inheritance() {
    let src = r#"
class BaseCollection {
    *[Symbol.iterator]() {
        yield 1; yield 2;
    }
}
class DerivedCollection extends BaseCollection {}

console.log([...new DerivedCollection()].join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2"]);
}

#[test]
fn test_js_symbol_iterator_reuse_iterator_object() {
    let src = r#"
const iter = [1, 2][Symbol.iterator]();
const selfIter = {
    [Symbol.iterator]() { return iter; }
};
console.log([...selfIter].join(",") + "|" + [...selfIter].length); // Second iteration is empty because iter is exhausted!
"#;
    assert_eq!(run_js(src), vec!["1,2|0"]);
}

#[test]
fn test_js_symbol_async_iterator_reject_in_next_propagates() {
    let src = r#"
const failAsyncIter = {
    [Symbol.asyncIterator]() {
        return {
            next() { return Promise.reject("AsyncNextFailed"); }
        };
    }
};
(async () => {
    try {
        for await (const _ of failAsyncIter);
    } catch (e) {
        console.log("Caught: " + e);
    }
})();
"#;
    assert_eq!(run_js(src), vec!["Caught: AsyncNextFailed"]);
}

#[test]
fn test_js_symbol_iterator_array_from_custom_iterable() {
    let src = r#"
const obj = {
    [Symbol.iterator]: function*() { yield 5; yield 10; }
};
console.log(Array.from(obj, x => x * 2).join(","));
"#;
    assert_eq!(run_js(src), vec!["10,20"]);
}

#[test]
fn test_js_symbol_iterator_map_keys_values_entries() {
    let src = r#"
const map = new Map([["a", 1]]);
console.log(typeof map[Symbol.iterator] === "function");
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_symbol_iterator_string_primitive() {
    let src = r#"
const str = "A";
const iter = str[Symbol.iterator]();
console.log(iter.next().value + "|done=" + iter.next().done);
"#;
    assert_eq!(run_js(src), vec!["A|done=true"]);
}

#[test]
fn test_js_symbol_iterator_in_yield_star_delegation() {
    let src = r#"
const custom = {
    *[Symbol.iterator]() { yield "X"; yield "Y"; }
};
function* outer() {
    yield* custom;
}
console.log([...outer()].join("-"));
"#;
    assert_eq!(run_js(src), vec!["X-Y"]);
}

#[test]
fn test_js_symbol_async_iterator_in_yield_star_async_generator() {
    let src = r#"
const asyncObj = {
    async *[Symbol.asyncIterator]() { yield "A1"; }
};
async function* outer() {
    yield* asyncObj;
}
(async () => {
    for await (const val of outer()) console.log(val);
})();
"#;
    assert_eq!(run_js(src), vec!["A1"]);
}

#[test]
fn test_js_symbol_iterator_well_known_symbol_identity() {
    let src = r#"
console.log(typeof Symbol.iterator === "symbol" && typeof Symbol.asyncIterator === "symbol");
"#;
    assert_eq!(run_js(src), vec!["true"]);
}
