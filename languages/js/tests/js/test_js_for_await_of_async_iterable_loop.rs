use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: for-await-of Async Iterable Loop Semantics
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_for_await_of_sync_array() {
    let src = r#"
(async () => {
    const results = [];
    for await (const x of [10, 20, 30]) {
        results.push(x * 2);
    }
    console.log(results.join(","));
})();
"#;
    assert_eq!(run_js(src), vec!["20,40,60"]);
}

#[test]
fn test_js_for_await_of_array_of_promises() {
    let src = r#"
(async () => {
    const promises = [Promise.resolve("A"), Promise.resolve("B")];
    const items = [];
    for await (const item of promises) {
        items.push(item);
    }
    console.log(items.join(","));
})();
"#;
    assert_eq!(run_js(src), vec!["A,B"]);
}

#[test]
fn test_js_for_await_of_custom_async_iterable() {
    let src = r#"
const customAsyncIterable = {
    [Symbol.asyncIterator]() {
        let count = 0;
        return {
            async next() {
                if (count < 3) {
                    return { value: ++count, done: false };
                }
                return { value: undefined, done: true };
            }
        };
    }
};
(async () => {
    const vals = [];
    for await (const v of customAsyncIterable) {
        vals.push(v);
    }
    console.log(vals.join(","));
})();
"#;
    assert_eq!(run_js(src), vec!["1,2,3"]);
}

#[test]
fn test_js_for_await_of_break_calls_return_method() {
    let src = r#"
let returnCalled = false;
const customAsyncIterable = {
    [Symbol.asyncIterator]() {
        return {
            async next() { return { value: 1, done: false }; },
            async return() { returnCalled = true; return { done: true }; }
        };
    }
};
(async () => {
    for await (const x of customAsyncIterable) {
        break;
    }
    console.log("Return Method Called: " + returnCalled);
})();
"#;
    assert_eq!(run_js(src), vec!["Return Method Called: true"]);
}

#[test]
fn test_js_for_await_of_throw_calls_return_method() {
    let src = r#"
let returnCalled = false;
const customAsyncIterable = {
    [Symbol.asyncIterator]() {
        return {
            async next() { return { value: 1, done: false }; },
            async return() { returnCalled = true; return { done: true }; }
        };
    }
};
(async () => {
    try {
        for await (const x of customAsyncIterable) {
            throw new Error("LoopCrash");
        }
    } catch (e) {
        console.log(e.message + "|ReturnCalled=" + returnCalled);
    }
})();
"#;
    assert_eq!(run_js(src), vec!["LoopCrash|ReturnCalled=true"]);
}

#[test]
fn test_js_for_await_of_non_async_iterable_throws_typeerror() {
    let src = r#"
(async () => {
    try {
        for await (const x of 12345) {}
    } catch (e) {
        console.log("For-Await-Of Non-Iterable TypeError");
    }
})();
"#;
    assert_eq!(run_js(src), vec!["For-Await-Of Non-Iterable TypeError"]);
}

#[test]
fn test_js_for_await_of_destructured_bindings() {
    let src = r#"
(async () => {
    const pairs = [Promise.resolve([1, "a"]), Promise.resolve([2, "b"])];
    const log = [];
    for await (const [id, label] of pairs) {
        log.push(`${id}:${label}`);
    }
    console.log(log.join(","));
})();
"#;
    assert_eq!(run_js(src), vec!["1:a,2:b"]);
}

#[test]
fn test_js_for_await_of_let_scoping_per_iteration() {
    let src = r#"
(async () => {
    const fns = [];
    for await (const i of [1, 2, 3]) {
        fns.push(() => i);
    }
    console.log(fns.map(f => f()).join(","));
})();
"#;
    assert_eq!(run_js(src), vec!["1,2,3"]);
}

#[test]
fn test_js_for_await_of_set_collection() {
    let src = r#"
(async () => {
    const set = new Set(["X", "Y"]);
    const items = [];
    for await (const item of set) {
        items.push(item);
    }
    console.log(items.join(","));
})();
"#;
    assert_eq!(run_js(src), vec!["X,Y"]);
}

#[test]
fn test_js_for_await_of_map_collection_entries() {
    let src = r#"
(async () => {
    const map = new Map([["k1", 100], ["k2", 200]]);
    const log = [];
    for await (const [k, v] of map) {
        log.push(`${k}=${v}`);
    }
    console.log(log.join(","));
})();
"#;
    assert_eq!(run_js(src), vec!["k1=100,k2=200"]);
}

#[test]
fn test_js_for_await_of_async_generator_consumption() {
    let src = r#"
async function* stream() {
    yield "Chunk1";
    yield "Chunk2";
}
(async () => {
    const chunks = [];
    for await (const chunk of stream()) {
        chunks.push(chunk);
    }
    console.log(chunks.join("|"));
})();
"#;
    assert_eq!(run_js(src), vec!["Chunk1|Chunk2"]);
}

#[test]
fn test_js_for_await_of_rejection_during_iteration_halts_loop() {
    let src = r#"
(async () => {
    const items = [Promise.resolve(1), Promise.reject("IterFail"), Promise.resolve(3)];
    const log = [];
    try {
        for await (const item of items) {
            log.push(item);
        }
    } catch (e) {
        log.push("Caught:" + e);
    }
    console.log(log.join(","));
})();
"#;
    assert_eq!(run_js(src), vec!["1,Caught:IterFail"]);
}

#[test]
fn test_js_for_await_of_continue_statement() {
    let src = r#"
(async () => {
    const numbers = [1, 2, 3, 4, 5];
    const evens = [];
    for await (const n of numbers) {
        if (n % 2 !== 0) continue;
        evens.push(n);
    }
    console.log(evens.join(","));
})();
"#;
    assert_eq!(run_js(src), vec!["2,4"]);
}

#[test]
fn test_js_for_await_of_string_code_points() {
    let src = r#"
(async () => {
    const chars = [];
    for await (const char of "JS") {
        chars.push(char);
    }
    console.log(chars.join("-"));
})();
"#;
    assert_eq!(run_js(src), vec!["J-S"]);
}

#[test]
fn test_js_for_await_of_typed_array() {
    let src = r#"
(async () => {
    const u8 = new Uint8Array([5, 10, 15]);
    const values = [];
    for await (const v of u8) {
        values.push(v);
    }
    console.log(values.join(","));
})();
"#;
    assert_eq!(run_js(src), vec!["5,10,15"]);
}

#[test]
fn test_js_for_await_of_fallback_to_sync_iterator() {
    let src = r#"
const syncIterable = {
    [Symbol.iterator]() {
        let i = 0;
        return {
            next() {
                return i < 2 ? { value: ++i, done: false } : { done: true };
            }
        };
    }
};
(async () => {
    const res = [];
    for await (const x of syncIterable) {
        res.push(x);
    }
    console.log(res.join(","));
})();
"#;
    assert_eq!(run_js(src), vec!["1,2"]);
}

#[test]
fn test_js_for_await_of_empty_sequence() {
    let src = r#"
(async () => {
    let count = 0;
    for await (const x of []) {
        count++;
    }
    console.log("Count: " + count);
})();
"#;
    assert_eq!(run_js(src), vec!["Count: 0"]);
}

#[test]
fn test_js_for_await_of_label_break() {
    let src = r#"
(async () => {
    const log = [];
    outer: for await (const i of [1, 2]) {
        for await (const j of [10, 20]) {
            if (i === 1 && j === 20) break outer;
            log.push(`${i}:${j}`);
        }
    }
    console.log(log.join(","));
})();
"#;
    assert_eq!(run_js(src), vec!["1:10"]);
}

#[test]
fn test_js_for_await_of_nested_async_generators() {
    let src = r#"
async function* inner() { yield "A"; yield "B"; }
async function* outer() {
    for await (const item of inner()) {
        yield `Wrapped_${item}`;
    }
}
(async () => {
    const results = [];
    for await (const v of outer()) results.push(v);
    console.log(results.join(","));
})();
"#;
    assert_eq!(run_js(src), vec!["Wrapped_A,Wrapped_B"]);
}

#[test]
fn test_js_for_await_of_return_in_function_body() {
    let src = r#"
async function findFirstEven(numbers) {
    for await (const n of numbers) {
        if (n % 2 === 0) return n;
    }
    return null;
}
findFirstEven([1, 3, 6, 7]).then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["6"]);
}

#[test]
fn test_js_for_await_of_non_object_next_throws_typeerror() {
    let src = r#"
(async () => {
    const invalid = {
        [Symbol.asyncIterator]() {
            return {
                next() {
                    return 123; // invalid: next must return an object
                }
            };
        }
    };
    try {
        for await (const _ of invalid) {}
    } catch (e) {
        console.log(e instanceof TypeError ? "TypeError" : e.constructor.name);
    }
})();
"#;
    assert_eq!(run_js(src), vec!["TypeError"]);
}

#[test]
fn test_js_for_await_of_continue_does_not_call_return() {
    let src = r#"
(async () => {
    let returnCalls = 0;
    const iterable = {
        [Symbol.asyncIterator]() {
            let i = 0;
            return {
                async next() {
                    return i < 4 ? { value: ++i, done: false } : { done: true };
                },
                async return() {
                    returnCalls++;
                    return { done: true };
                }
            };
        }
    };

    const seen = [];
    for await (const n of iterable) {
        if (n % 2 === 0) continue;
        seen.push(n);
    }
    console.log(seen.join(","));
    console.log(returnCalls);
})();
"#;
    assert_eq!(run_js(src), vec!["1,3", "0"]);
}

#[test]
fn test_js_for_await_of_async_generator_yields_rejected_promise() {
    let src = r#"
(async () => {
    async function* gen() {
        yield Promise.reject("GenReject");
    }
    try {
        for await (const x of gen()) {}
    } catch(e) {
        console.log(e);
    }
})();
"#;
    assert_eq!(run_js(src), vec!["GenReject"]);
}

#[test]
fn for_await_of_prefers_symbol_async_iterator_over_sync() {
    let src = r#"
const dual = {
    [Symbol.iterator]() {
        return { next() { return { value: "sync", done: false }; } };
    },
    [Symbol.asyncIterator]() {
        return { async next() { return { value: "async", done: false }; } };
    }
};
(async () => {
    for await (const x of dual) {
        console.log(x);
        break;
    }
})();
"#;
    assert_eq!(run_js(src), vec!["async"]);
}
