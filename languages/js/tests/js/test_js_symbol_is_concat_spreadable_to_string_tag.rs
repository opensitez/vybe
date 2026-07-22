use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `Symbol.isConcatSpreadable` & `Symbol.toStringTag` Metaprogramming Hooks
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_symbol_is_concat_spreadable_array_flattening_disabled() {
    let src = r#"
const arr1 = [1, 2];
const arr2 = [3, 4];
arr2[Symbol.isConcatSpreadable] = false; // Prevents concat from flattening arr2!

const res = arr1.concat(arr2);
console.log(res.length + "|" + Array.isArray(res[2]));
"#;
    assert_eq!(run_js(src), vec!["3|true"]);
}

#[test]
fn test_js_symbol_is_concat_spreadable_array_like_object_flattening_enabled() {
    let src = r#"
const arrayLike = {
    0: "A", 1: "B", length: 2,
    [Symbol.isConcatSpreadable]: true
};
const res = [1].concat(arrayLike);
console.log(res.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,A,B"]);
}

#[test]
fn test_js_symbol_tostringtag_custom_class() {
    let src = r#"
class Validator {
    get [Symbol.toStringTag]() {
        return "CustomValidator";
    }
}
const v = new Validator();
console.log(Object.prototype.toString.call(v));
"#;
    assert_eq!(run_js(src), vec!["[object CustomValidator]"]);
}

#[test]
fn test_js_symbol_tostringtag_object_literal() {
    let src = r#"
const moduleObj = {
    [Symbol.toStringTag]: "MyModule"
};
console.log(Object.prototype.toString.call(moduleObj));
"#;
    assert_eq!(run_js(src), vec!["[object MyModule]"]);
}

#[test]
fn test_js_symbol_is_concat_spreadable_truthy_falsy_coercion() {
    let src = r#"
const arr = [10, 20];
arr[Symbol.isConcatSpreadable] = 0; // Falsy -> not spreadable
const res1 = [0].concat(arr);

arr[Symbol.isConcatSpreadable] = 1; // Truthy -> spreadable
const res2 = [0].concat(arr);
console.log(res1.length + "|" + res2.length);
"#;
    assert_eq!(run_js(src), vec!["2|3"]);
}

#[test]
fn test_js_symbol_tostringtag_builtin_objects() {
    let src = r#"
console.log(`${Object.prototype.toString.call(new Map())}:${Object.prototype.toString.call(new Set())}:${Object.prototype.toString.call(Promise.resolve())}`);
"#;
    assert_eq!(
        run_js(src),
        vec!["[object Map]:[object Set]:[object Promise]"]
    );
}

#[test]
fn test_js_symbol_tostringtag_generator_function() {
    let src = r#"
function* myGenerator() {}
console.log(Object.prototype.toString.call(myGenerator()));
"#;
    assert_eq!(run_js(src), vec!["[object Generator]"]);
}

#[test]
fn test_js_symbol_tostringtag_async_generator_function() {
    let src = r#"
async function* myAsyncGen() {}
console.log(Object.prototype.toString.call(myAsyncGen()));
"#;
    assert_eq!(run_js(src), vec!["[object AsyncGenerator]"]);
}

#[test]
fn test_js_symbol_is_concat_spreadable_defaults() {
    let src = r#"
const array = [1];
const plainObj = { 0: "a", length: 1 };
const res = [0].concat(array, plainObj);
console.log(res.length + "|" + Array.isArray(res[1]) + "|" + (typeof res[2]));
"#;
    assert_eq!(run_js(src), vec!["3|false|object"]); // Array is spreadable by default, plain object is not
}

#[test]
fn test_js_symbol_tostringtag_non_string_ignored() {
    let src = r#"
const obj = {
    [Symbol.toStringTag]: 12345 // Non-string is ignored, falls back to default Object tag!
};
console.log(Object.prototype.toString.call(obj));
"#;
    assert_eq!(run_js(src), vec!["[object Object]"]);
}

#[test]
fn test_js_symbol_is_concat_spreadable_null_prototype_array_like() {
    let src = r#"
const arrayLike = Object.create(null);
arrayLike[0] = "x";
arrayLike.length = 1;
arrayLike[Symbol.isConcatSpreadable] = true;

const res = [1].concat(arrayLike);
console.log(res.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,x"]);
}

#[test]
fn test_js_symbol_tostringtag_math_and_json_objects() {
    let src = r#"
console.log(`${Object.prototype.toString.call(Math)}:${Object.prototype.toString.call(JSON)}`);
"#;
    assert_eq!(run_js(src), vec!["[object Math]:[object JSON]"]);
}

#[test]
fn test_js_symbol_tostringtag_globalthis_object() {
    let src = r#"
console.log(Object.prototype.toString.call(globalThis));
"#;
    assert_eq!(run_js(src), vec!["[object global]"]);
}

#[test]
fn test_js_symbol_is_concat_spreadable_subclass_array() {
    let src = r#"
class SubArray extends Array {}
const sa = new SubArray(1, 2);
sa[Symbol.isConcatSpreadable] = false;
const res = [0].concat(sa);
console.log(res.length + "|" + (res[1] instanceof SubArray));
"#;
    assert_eq!(run_js(src), vec!["2|true"]);
}

#[test]
fn test_js_symbol_tostringtag_symbol_prototype() {
    let src = r#"
const s = Symbol("test");
console.log(Object.prototype.toString.call(s));
"#;
    assert_eq!(run_js(src), vec!["[object Symbol]"]);
}

#[test]
fn test_js_symbol_tostringtag_bigint_prototype() {
    let src = r#"
const b = 100n;
console.log(Object.prototype.toString.call(b));
"#;
    assert_eq!(run_js(src), vec!["[object BigInt]"]);
}

#[test]
fn test_js_symbol_is_concat_spreadable_sparse_array_holes_preserved() {
    let src = r#"
const sparse = [1, , 3];
sparse[Symbol.isConcatSpreadable] = true;
const res = [0].concat(sparse);
console.log(res.length + "|hasHole=" + !(2 in res));
"#;
    assert_eq!(run_js(src), vec!["4|hasHole=true"]);
}

#[test]
fn test_js_symbol_tostringtag_prototype_property_descriptor() {
    let src = r#"
const desc = Object.getOwnPropertyDescriptor(Symbol, "toStringTag");
console.log(desc.writable + "|" + desc.enumerable + "|" + desc.configurable);
"#;
    assert_eq!(run_js(src), vec!["false|false|false"]);
}

#[test]
fn test_js_symbol_is_concat_spreadable_property_descriptor() {
    let src = r#"
const desc = Object.getOwnPropertyDescriptor(Symbol, "isConcatSpreadable");
console.log(desc.writable + "|" + desc.enumerable + "|" + desc.configurable);
"#;
    assert_eq!(run_js(src), vec!["false|false|false"]);
}

#[test]
fn test_js_symbol_tostringtag_well_known_symbols_exist() {
    let src = r#"
console.log(typeof Symbol.isConcatSpreadable === "symbol" && typeof Symbol.toStringTag === "symbol");
"#;
    assert_eq!(run_js(src), vec!["true"]);
}
