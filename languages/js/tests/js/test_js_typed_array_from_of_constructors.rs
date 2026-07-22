use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: TypedArray `.from()` & `.of()` Static Constructors
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_typedarray_of_basic_factory() {
    let src = r#"
const u8 = Uint8Array.of(10, 20, 30);
console.log(u8.length + "|" + u8.join(","));
"#;
    assert_eq!(run_js(src), vec!["3|10,20,30"]);
}

#[test]
fn test_js_typedarray_from_array_factory() {
    let src = r#"
const i32 = Int32Array.from([100, 200, 300]);
console.log(i32.length + "|" + i32.join(","));
"#;
    assert_eq!(run_js(src), vec!["3|100,200,300"]);
}

#[test]
fn test_js_typedarray_from_mapping_function() {
    let src = r#"
const u8 = Uint8Array.from([1, 2, 3], x => x * 10);
console.log(u8.join(","));
"#;
    assert_eq!(run_js(src), vec!["10,20,30"]);
}

#[test]
fn test_js_typedarray_from_this_arg_binding() {
    let src = r#"
const ctx = { scale: 5 };
const res = Uint8Array.from([1, 2], function(x) { return x * this.scale; }, ctx);
console.log(res.join(","));
"#;
    assert_eq!(run_js(src), vec!["5,10"]);
}

#[test]
fn test_js_typedarray_from_string_iterable() {
    let src = r#"
const u8 = Uint8Array.from("123", x => Number(x));
console.log(u8.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2,3"]);
}

#[test]
fn test_js_typedarray_from_set_iterable() {
    let src = r#"
const set = new Set([5, 10, 15]);
const i16 = Int16Array.from(set);
console.log(i16.join(","));
"#;
    assert_eq!(run_js(src), vec!["5,10,15"]);
}

#[test]
fn test_js_typedarray_from_generator_iterable() {
    let src = r#"
function* gen() { yield 2; yield 4; yield 6; }
const f64 = Float64Array.from(gen());
console.log(f64.join(","));
"#;
    assert_eq!(run_js(src), vec!["2,4,6"]);
}

#[test]
fn test_js_typedarray_from_array_like_object() {
    let src = r#"
const arrayLike = { 0: 10, 1: 20, length: 2 };
const u8 = Uint8Array.from(arrayLike);
console.log(u8.join(","));
"#;
    assert_eq!(run_js(src), vec!["10,20"]);
}

#[test]
fn test_js_typedarray_of_element_coercion() {
    let src = r#"
const u8 = Uint8Array.of("50", 256, true); // "50"->50, 256->0, true->1
console.log(u8.join(","));
"#;
    assert_eq!(run_js(src), vec!["50,0,1"]);
}

#[test]
fn test_js_typedarray_from_typedarray_reencoding() {
    let src = r#"
const f32 = new Float32Array([1.5, 2.5]);
const u8 = Uint8Array.from(f32);
console.log(u8.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2"]);
}

#[test]
fn test_js_typedarray_from_mapping_function_index_arg() {
    let src = r#"
const res = Uint8Array.from([0, 0, 0], (val, idx) => idx + 1);
console.log(res.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2,3"]);
}

#[test]
fn test_js_typedarray_from_null_iterable_throws_typeerror() {
    let src = r#"
try {
    Uint8Array.from(null);
} catch (e) {
    console.log("Uint8Array.from Null TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Uint8Array.from Null TypeError"]);
}

#[test]
fn test_js_typedarray_from_non_callable_mapfn_throws_typeerror() {
    let src = r#"
try {
    Uint8Array.from([1, 2], "not_a_function");
} catch (e) {
    console.log("Uint8Array.from Non-Callable MapFn TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Uint8Array.from Non-Callable MapFn TypeError"]
    );
}

#[test]
fn test_js_typedarray_of_empty_args_returns_zero_length() {
    let src = r#"
const empty = Int32Array.of();
console.log(empty.length);
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_typedarray_from_empty_iterable() {
    let src = r#"
const empty = Float64Array.from([]);
console.log(empty.length);
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_typedarray_from_subclass_constructor_inheritance() {
    let src = r#"
class CustomUint8 extends Uint8Array {}
const cu8 = CustomUint8.from([1, 2, 3]);
console.log(cu8.join(",") + "|isCustom=" + (cu8 instanceof CustomUint8));
"#;
    assert_eq!(run_js(src), vec!["1,2,3|isCustom=true"]);
}

#[test]
fn test_js_typedarray_of_subclass_constructor_inheritance() {
    let src = r#"
class CustomUint8 extends Uint8Array {}
const cu8 = CustomUint8.of(10, 20);
console.log(cu8.join(",") + "|isCustom=" + (cu8 instanceof CustomUint8));
"#;
    assert_eq!(run_js(src), vec!["10,20|isCustom=true"]);
}

#[test]
fn test_js_typedarray_from_bigint64_view() {
    let src = r#"
const big = BigInt64Array.from([100n, 200n]);
console.log(big.length + "|" + big[0].toString());
"#;
    assert_eq!(run_js(src), vec!["2|100"]);
}

#[test]
fn test_js_typedarray_of_bigint64_view() {
    let src = r#"
const big = BigInt64Array.of(500n, 600n);
console.log(big[0].toString() + ":" + big[1].toString());
"#;
    assert_eq!(run_js(src), vec!["500:600"]);
}

#[test]
fn test_js_typedarray_from_sparse_array_holes_mapped_to_undefined() {
    let src = r#"
const sparse = [1, , 3];
const res = Uint8Array.from(sparse, x => x === undefined ? 99 : x);
console.log(res.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,99,3"]);
}
