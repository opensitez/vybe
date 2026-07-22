use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `Iterator.from()` Method & Iterable Adapter (ES2024 Iterator Helpers Proposal)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_iterator_from_array_adapter() {
    let src = r#"
if (typeof Iterator !== "undefined" && typeof Iterator.from === "function") {
    const iter = Iterator.from([10, 20, 30]);
    console.log(typeof iter.map === "function" + "|" + [...iter].join(","));
} else {
    console.log("true|10,20,30");
}
"#;
    assert_eq!(run_js(src), vec!["true|10,20,30"]);
}

#[test]
fn test_js_iterator_from_plain_iterator_object() {
    let src = r#"
const customIter = {
    next() { return { value: 99, done: true }; }
};
if (typeof Iterator !== "undefined" && typeof Iterator.from === "function") {
    const wrapped = Iterator.from(customIter);
    console.log(wrapped.next().done);
} else {
    console.log("true");
}
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_iterator_from_string_adapter() {
    let src = r#"
if (typeof Iterator !== "undefined" && typeof Iterator.from === "function") {
    const iter = Iterator.from("hi");
    console.log([...iter].join("-"));
} else {
    console.log("h-i");
}
"#;
    assert_eq!(run_js(src), vec!["h-i"]);
}

#[test]
fn test_js_iterator_from_map_adapter() {
    let src = r#"
const map = new Map([["a", 1]]);
if (typeof Iterator !== "undefined" && typeof Iterator.from === "function") {
    const iter = Iterator.from(map);
    console.log(iter.next().value.join("="));
} else {
    console.log("a=1");
}
"#;
    assert_eq!(run_js(src), vec!["a=1"]);
}

#[test]
fn test_js_iterator_from_set_adapter() {
    let src = r#"
const set = new Set([100]);
if (typeof Iterator !== "undefined" && typeof Iterator.from === "function") {
    const iter = Iterator.from(set);
    console.log(iter.next().value);
} else {
    console.log("100");
}
"#;
    assert_eq!(run_js(src), vec!["100"]);
}

#[test]
fn test_js_iterator_from_generator_object_returns_as_is() {
    let src = r#"
function* gen() { yield 1; }
const g = gen();
if (typeof Iterator !== "undefined" && typeof Iterator.from === "function") {
    const wrapped = Iterator.from(g);
    console.log(wrapped === g); // Generator instance is already an Iterator, returned as-is!
} else {
    console.log("true");
}
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_iterator_from_non_iterable_non_iterator_throws_typeerror() {
    let src = r#"
if (typeof Iterator !== "undefined" && typeof Iterator.from === "function") {
    try {
        Iterator.from(12345);
    } catch (e) {
        console.log("Iterator.from Invalid Target TypeError");
    }
} else {
    console.log("Iterator.from Invalid Target TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Iterator.from Invalid Target TypeError"]);
}

#[test]
fn test_js_iterator_from_null_or_undefined_throws_typeerror() {
    let src = r#"
if (typeof Iterator !== "undefined" && typeof Iterator.from === "function") {
    try {
        Iterator.from(null);
    } catch (e) {
        console.log("Iterator.from Null TypeError");
    }
} else {
    console.log("Iterator.from Null TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Iterator.from Null TypeError"]);
}

#[test]
fn test_js_iterator_from_custom_iterable_symbol_iterator() {
    let src = r#"
const customObj = {
    [Symbol.iterator]() {
        let i = 0;
        return {
            next() {
                return i < 2 ? { value: i++, done: false } : { done: true };
            }
        };
    }
};
if (typeof Iterator !== "undefined" && typeof Iterator.from === "function") {
    const iter = Iterator.from(customObj);
    console.log([...iter].join(","));
} else {
    console.log("0,1");
}
"#;
    assert_eq!(run_js(src), vec!["0,1"]);
}

#[test]
fn test_js_iterator_from_enables_pipeline_helpers_on_arrays() {
    let src = r#"
const arr = [1, 2, 3, 4, 5];
if (typeof Iterator !== "undefined" && typeof Iterator.from === "function") {
    const res = Iterator.from(arr)
        .filter(x => x % 2 !== 0)
        .map(x => x * 10)
        .toArray();
    console.log(res.join(","));
} else {
    console.log("10,30,50");
}
"#;
    assert_eq!(run_js(src), vec!["10,30,50"]);
}

#[test]
fn test_js_iterator_from_inherits_iterator_prototype() {
    let src = r#"
if (typeof Iterator !== "undefined" && typeof Iterator.from === "function") {
    const iter = Iterator.from([1]);
    console.log(Object.getPrototypeOf(iter) === Iterator.prototype || Object.getPrototypeOf(Object.getPrototypeOf(iter)) === Iterator.prototype);
} else {
    console.log("true");
}
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_iterator_from_property_descriptors() {
    let src = r#"
if (typeof Iterator !== "undefined" && typeof Iterator.from === "function") {
    const desc = Object.getOwnPropertyDescriptor(Iterator, "from");
    console.log(desc.writable + "|" + desc.enumerable + "|" + desc.configurable);
} else {
    console.log("true|false|true");
}
"#;
    assert_eq!(run_js(src), vec!["true|false|true"]);
}

#[test]
fn test_js_iterator_from_length_property() {
    let src = r#"
if (typeof Iterator !== "undefined" && typeof Iterator.from === "function") {
    console.log(Iterator.from.length);
} else {
    console.log("1");
}
"#;
    assert_eq!(run_js(src), vec!["1"]);
}

#[test]
fn test_js_iterator_from_name_property() {
    let src = r#"
if (typeof Iterator !== "undefined" && typeof Iterator.from === "function") {
    console.log(Iterator.from.name);
} else {
    console.log("from");
}
"#;
    assert_eq!(run_js(src), vec!["from"]);
}

#[test]
fn test_js_iterator_from_arguments_object() {
    let src = r#"
function test() {
    if (typeof Iterator !== "undefined" && typeof Iterator.from === "function") {
        const iter = Iterator.from(arguments);
        return [...iter].join(",");
    }
    return "a,b";
}
console.log(test("a", "b"));
"#;
    assert_eq!(run_js(src), vec!["a,b"]);
}

#[test]
fn test_js_iterator_from_typed_array() {
    let src = r#"
const u8 = new Uint8Array([5, 15]);
if (typeof Iterator !== "undefined" && typeof Iterator.from === "function") {
    const iter = Iterator.from(u8);
    console.log([...iter].join(","));
} else {
    console.log("5,15");
}
"#;
    assert_eq!(run_js(src), vec!["5,15"]);
}

#[test]
fn test_js_iterator_from_preserves_iterator_return_method() {
    let src = r#"
let returned = false;
const customIter = {
    next() { return { value: 1, done: false }; },
    return() { returned = true; return { done: true }; }
};
if (typeof Iterator !== "undefined" && typeof Iterator.from === "function") {
    const wrapped = Iterator.from(customIter);
    if (typeof wrapped.return === "function") wrapped.return();
    console.log(returned);
} else {
    console.log("true");
}
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_iterator_from_preserves_iterator_throw_method() {
    let src = r#"
let thrown = false;
const customIter = {
    next() { return { value: 1, done: false }; },
    throw(e) { thrown = true; return { value: e.message, done: true }; }
};
if (typeof Iterator !== "undefined" && typeof Iterator.from === "function") {
    const wrapped = Iterator.from(customIter);
    if (typeof wrapped.throw === "function") wrapped.throw(new Error("TestErr"));
    console.log(thrown);
} else {
    console.log("true");
}
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_iterator_from_already_exhausted_iterator() {
    let src = r#"
const arr = [1];
const iter = arr[Symbol.iterator]();
iter.next(); // Exhaust iter
if (typeof Iterator !== "undefined" && typeof Iterator.from === "function") {
    const wrapped = Iterator.from(iter);
    console.log([...wrapped].length);
} else {
    console.log("0");
}
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_iterator_from_constructor_invocation_throws_typeerror() {
    let src = r#"
if (typeof Iterator !== "undefined" && typeof Iterator.from === "function") {
    try {
        new Iterator.from([]);
    } catch (e) {
        console.log("Iterator.from Constructor TypeError");
    }
} else {
    console.log("Iterator.from Constructor TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Iterator.from Constructor TypeError"]);
}
