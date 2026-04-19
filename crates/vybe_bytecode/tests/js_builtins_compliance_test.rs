//! Phase B6 — `wasm:js-*` builtin byte-level compliance tests.
//!
//! These tests verify the **emitter side** of the JS builtin import
//! surface: every declared import gets a valid signature, imports
//! appear in emitted `.wasm` under the right module name, signatures
//! match the conventions pinned in `JS_BUILTIN_CONVENTIONS.md`.
//!
//! Behavioral tests (handler logic + end-to-end execution via
//! CALL_IMPORT) live in `vybe_host/tests/js_builtins_behavior_test.rs`
//! since `vybe_bytecode` can't depend on `vybe_host`.
//!
//! See `dynamicruntime_support.md` Phase B6.

use vybe_bytecode::wasm::{
    js_array_builtins, js_arraybuffer_builtins, js_map_builtins,
    js_object_builtins, js_set_builtins, js_typedarray_builtins,
    js_weakmap_builtins,
};

// ──────────────────────────────────────────────────────────────────────
// Contract 1: every declared import produces a valid signature
// ──────────────────────────────────────────────────────────────────────

#[test]
fn every_array_import_has_signature() {
    for name in js_array_builtins::IMPORTS {
        let mut buf = Vec::new();
        let ok = js_array_builtins::write_signature(&mut buf, name);
        assert!(ok, "wasm:js-array.{} declared but no signature", name);
        assert!(!buf.is_empty(),
            "wasm:js-array.{} signature is empty", name);
    }
}

#[test]
fn every_object_import_has_signature() {
    for name in js_object_builtins::IMPORTS {
        let mut buf = Vec::new();
        let ok = js_object_builtins::write_signature(&mut buf, name);
        assert!(ok, "wasm:js-object.{} declared but no signature", name);
        assert!(!buf.is_empty());
    }
}

#[test]
fn every_map_import_has_signature() {
    for name in js_map_builtins::IMPORTS {
        let mut buf = Vec::new();
        let ok = js_map_builtins::write_signature(&mut buf, name);
        assert!(ok, "wasm:js-map.{} declared but no signature", name);
        assert!(!buf.is_empty());
    }
}

#[test]
fn every_set_import_has_signature() {
    for name in js_set_builtins::IMPORTS {
        let mut buf = Vec::new();
        let ok = js_set_builtins::write_signature(&mut buf, name);
        assert!(ok, "wasm:js-set.{} declared but no signature", name);
        assert!(!buf.is_empty());
    }
}

#[test]
fn every_weakmap_and_weakset_import_has_signature() {
    for name in js_weakmap_builtins::WEAKMAP_IMPORTS {
        let mut buf = Vec::new();
        let ok = js_weakmap_builtins::write_weakmap_signature(&mut buf, name);
        assert!(ok, "wasm:js-weakmap.{} declared but no signature", name);
    }
    for name in js_weakmap_builtins::WEAKSET_IMPORTS {
        let mut buf = Vec::new();
        let ok = js_weakmap_builtins::write_weakset_signature(&mut buf, name);
        assert!(ok, "wasm:js-weakset.{} declared but no signature", name);
    }
}

#[test]
fn every_arraybuffer_import_has_signature() {
    for name in js_arraybuffer_builtins::ARRAYBUFFER_IMPORTS {
        let mut buf = Vec::new();
        let ok = js_arraybuffer_builtins::write_arraybuffer_signature(&mut buf, name);
        assert!(ok, "wasm:js-arraybuffer.{} declared but no signature", name);
    }
    for name in js_arraybuffer_builtins::SHAREDARRAYBUFFER_IMPORTS {
        let mut buf = Vec::new();
        let ok = js_arraybuffer_builtins::write_sharedarraybuffer_signature(&mut buf, name);
        assert!(ok, "wasm:js-sharedarraybuffer.{} declared but no signature", name);
    }
    for name in js_arraybuffer_builtins::DATAVIEW_IMPORTS {
        let mut buf = Vec::new();
        let ok = js_arraybuffer_builtins::write_dataview_signature(&mut buf, name);
        assert!(ok, "wasm:js-dataview.{} declared but no signature", name);
    }
}

#[test]
fn every_typedarray_variant_has_complete_signature_set() {
    use js_typedarray_builtins::{IMPORTS, VARIANTS};
    for variant in VARIANTS {
        for name in IMPORTS {
            let mut buf = Vec::new();
            let ok = js_typedarray_builtins::write_signature(&mut buf, *variant, name);
            assert!(ok,
                "{}.{} declared but no signature", variant.module(), name);
            assert!(!buf.is_empty());
        }
    }
}

// ──────────────────────────────────────────────────────────────────────
// Contract 2: no unknown import names succeed
// ──────────────────────────────────────────────────────────────────────

#[test]
fn rejects_unknown_import_names() {
    // An undeclared name must return false — catches typos in the
    // IMPORTS list vs the write_signature match arms.
    let mut buf = Vec::new();
    assert!(!js_array_builtins::write_signature(&mut buf, "nonsense_undeclared"));
    assert!(!js_object_builtins::write_signature(&mut buf, "nonsense_undeclared"));
    assert!(!js_map_builtins::write_signature(&mut buf, "nonsense_undeclared"));
    assert!(!js_set_builtins::write_signature(&mut buf, "nonsense_undeclared"));
    assert!(!js_arraybuffer_builtins::write_arraybuffer_signature(&mut buf, "xxx"));
    assert!(!js_arraybuffer_builtins::write_dataview_signature(&mut buf, "xxx"));
    assert!(!js_weakmap_builtins::write_weakmap_signature(&mut buf, "xxx"));
    assert!(!js_weakmap_builtins::write_weakset_signature(&mut buf, "xxx"));
    for variant in js_typedarray_builtins::VARIANTS {
        assert!(!js_typedarray_builtins::write_signature(&mut buf, *variant, "xxx"));
    }
}

// ──────────────────────────────────────────────────────────────────────
// Contract 3: module names match the JS-builtin namespace convention
// ──────────────────────────────────────────────────────────────────────

#[test]
fn module_names_follow_wasm_js_convention() {
    assert_eq!(js_array_builtins::MODULE, "wasm:js-array");
    assert_eq!(js_object_builtins::MODULE, "wasm:js-object");
    assert_eq!(js_map_builtins::MODULE, "wasm:js-map");
    assert_eq!(js_set_builtins::MODULE, "wasm:js-set");
    assert_eq!(js_weakmap_builtins::WEAKMAP_MODULE, "wasm:js-weakmap");
    assert_eq!(js_weakmap_builtins::WEAKSET_MODULE, "wasm:js-weakset");
    assert_eq!(js_arraybuffer_builtins::ARRAYBUFFER_MODULE, "wasm:js-arraybuffer");
    assert_eq!(js_arraybuffer_builtins::SHAREDARRAYBUFFER_MODULE,
        "wasm:js-sharedarraybuffer");
    assert_eq!(js_arraybuffer_builtins::DATAVIEW_MODULE, "wasm:js-dataview");

    // Each typed-array variant follows `wasm:js-<type>array` naming.
    for variant in js_typedarray_builtins::VARIANTS {
        assert!(variant.module().starts_with("wasm:js-"),
            "typed array module must start with wasm:js- prefix: {}",
            variant.module());
    }
}

// ──────────────────────────────────────────────────────────────────────
// Contract 4: no duplicate imports within a module
// ──────────────────────────────────────────────────────────────────────

#[test]
fn no_duplicate_imports_within_any_module() {
    fn assert_unique(module: &str, imports: &[&str]) {
        let mut seen = std::collections::HashSet::new();
        for name in imports {
            assert!(seen.insert(*name),
                "{}: duplicate import `{}`", module, name);
        }
    }

    assert_unique(js_array_builtins::MODULE, js_array_builtins::IMPORTS);
    assert_unique(js_object_builtins::MODULE, js_object_builtins::IMPORTS);
    assert_unique(js_map_builtins::MODULE, js_map_builtins::IMPORTS);
    assert_unique(js_set_builtins::MODULE, js_set_builtins::IMPORTS);
    assert_unique(js_weakmap_builtins::WEAKMAP_MODULE,
        js_weakmap_builtins::WEAKMAP_IMPORTS);
    assert_unique(js_weakmap_builtins::WEAKSET_MODULE,
        js_weakmap_builtins::WEAKSET_IMPORTS);
    assert_unique(js_arraybuffer_builtins::ARRAYBUFFER_MODULE,
        js_arraybuffer_builtins::ARRAYBUFFER_IMPORTS);
    assert_unique(js_arraybuffer_builtins::SHAREDARRAYBUFFER_MODULE,
        js_arraybuffer_builtins::SHAREDARRAYBUFFER_IMPORTS);
    assert_unique(js_arraybuffer_builtins::DATAVIEW_MODULE,
        js_arraybuffer_builtins::DATAVIEW_IMPORTS);
    assert_unique("typed-array method set", js_typedarray_builtins::IMPORTS);
}

// ──────────────────────────────────────────────────────────────────────
// Contract 5: signatures match the marshaling conventions
// ──────────────────────────────────────────────────────────────────────
//
// JS_BUILTIN_CONVENTIONS.md requires:
//   - Collection instances: externref (TYPE_EXTERNREF = 0x6F)
//   - Indices / sizes / booleans: i32 (TYPE_I32 = 0x7F)
//   - 64-bit integer values: i64 (TYPE_I64 = 0x7E)
//   - Float values: f64 (TYPE_F64 = 0x7C) — or f32 (0x7D) for Float32Array
//
// Every signature must be well-formed WASM: starts with a param-count
// LEB128, followed by param type bytes, then result-count LEB128 and
// result type bytes. No TYPE_FUNC prefix (the caller adds that).

fn decode_signature(buf: &[u8]) -> Option<(Vec<u8>, Vec<u8>)> {
    let (pc, pc_len) = decode_leb128_u32(buf);
    let param_end = pc_len + pc as usize;
    if param_end > buf.len() { return None; }
    let params: Vec<u8> = buf[pc_len..param_end].to_vec();
    let tail = &buf[param_end..];
    let (rc, rc_len) = decode_leb128_u32(tail);
    let result_end = rc_len + rc as usize;
    if result_end > tail.len() { return None; }
    let results: Vec<u8> = tail[rc_len..result_end].to_vec();
    Some((params, results))
}

fn decode_leb128_u32(bytes: &[u8]) -> (u32, usize) {
    let mut result: u32 = 0;
    let mut shift: u32 = 0;
    let mut pos = 0;
    loop {
        if pos >= bytes.len() { break; }
        let b = bytes[pos]; pos += 1;
        result |= ((b & 0x7F) as u32) << shift;
        if b & 0x80 == 0 { break; }
        shift += 7;
    }
    (result, pos)
}

#[test]
fn array_push_signature_is_externref_externref_to_i32() {
    // wasm:js-array.push: (arr: externref, v: externref) -> i32
    // Matches ECMA-262 §23.1.3.32 (returns the array's new length).
    let mut buf = Vec::new();
    assert!(js_array_builtins::write_signature(&mut buf, "push"));
    let (params, results) = decode_signature(&buf).expect("well-formed");
    assert_eq!(params, vec![0x6F, 0x6F], "params should be (externref, externref)");
    assert_eq!(results, vec![0x7F], "result should be i32 new_length");
}

#[test]
fn array_pop_signature_is_externref_to_externref() {
    // pop: (arr) -> popped_value (externref; undefined if empty)
    let mut buf = Vec::new();
    assert!(js_array_builtins::write_signature(&mut buf, "pop"));
    let (params, results) = decode_signature(&buf).unwrap();
    assert_eq!(params, vec![0x6F]);
    assert_eq!(results, vec![0x6F]);
}

#[test]
fn array_length_returns_i32() {
    let mut buf = Vec::new();
    assert!(js_array_builtins::write_signature(&mut buf, "length"));
    let (params, results) = decode_signature(&buf).unwrap();
    assert_eq!(params, vec![0x6F]);
    assert_eq!(results, vec![0x7F], "length must return i32 per convention");
}

#[test]
fn map_has_returns_i32_boolean() {
    // Booleans marshaled as i32 per convention.
    let mut buf = Vec::new();
    assert!(js_map_builtins::write_signature(&mut buf, "has"));
    let (params, results) = decode_signature(&buf).unwrap();
    assert_eq!(params, vec![0x6F, 0x6F]);
    assert_eq!(results, vec![0x7F], "has must return i32 0/1 per convention");
}

#[test]
fn dataview_get_float64_returns_f64_unboxed() {
    // Convention: float values are unboxed f64, not wrapped externref.
    let mut buf = Vec::new();
    assert!(js_arraybuffer_builtins::write_dataview_signature(&mut buf, "getFloat64"));
    let (params, results) = decode_signature(&buf).unwrap();
    assert_eq!(params, vec![0x6F, 0x7F, 0x7F],
        "getFloat64 params: (view, offset, littleEndian)");
    assert_eq!(results, vec![0x7C],
        "getFloat64 result must be f64 (0x7C), not boxed externref");
}

#[test]
fn dataview_get_biguint64_returns_i64() {
    let mut buf = Vec::new();
    assert!(js_arraybuffer_builtins::write_dataview_signature(&mut buf, "getBigUint64"));
    let (_, results) = decode_signature(&buf).unwrap();
    assert_eq!(results, vec![0x7E], "getBigUint64 result must be i64 (0x7E)");
}

#[test]
fn float32array_get_returns_f32() {
    let mut buf = Vec::new();
    assert!(js_typedarray_builtins::write_signature(
        &mut buf, js_typedarray_builtins::TypedElem::F32, "get"));
    let (params, results) = decode_signature(&buf).unwrap();
    assert_eq!(params, vec![0x6F, 0x7F]);
    assert_eq!(results, vec![0x7D],
        "Float32Array.get returns f32 (0x7D), matching BYTES_PER_ELEMENT=4");
}

#[test]
fn float64array_set_takes_f64_value() {
    let mut buf = Vec::new();
    assert!(js_typedarray_builtins::write_signature(
        &mut buf, js_typedarray_builtins::TypedElem::F64, "set"));
    let (params, results) = decode_signature(&buf).unwrap();
    assert_eq!(params, vec![0x6F, 0x7F, 0x7C],
        "Float64Array.set: (arr, index, f64_value)");
    assert_eq!(results, Vec::<u8>::new(), "set returns nothing");
}

#[test]
fn bigint64array_get_returns_i64() {
    let mut buf = Vec::new();
    assert!(js_typedarray_builtins::write_signature(
        &mut buf, js_typedarray_builtins::TypedElem::BigI64, "get"));
    let (_, results) = decode_signature(&buf).unwrap();
    assert_eq!(results, vec![0x7E],
        "BigInt64Array.get returns i64 (0x7E) per BigInt typed-array convention");
}

// ──────────────────────────────────────────────────────────────────────
// Contract 6: typed-array variant metadata is consistent
// ──────────────────────────────────────────────────────────────────────

#[test]
fn typed_array_variants_have_expected_bytes_per_element() {
    use js_typedarray_builtins::TypedElem;
    let expected: &[(TypedElem, u32)] = &[
        (TypedElem::I8, 1),
        (TypedElem::U8, 1),
        (TypedElem::U8Clamped, 1),
        (TypedElem::I16, 2),
        (TypedElem::U16, 2),
        (TypedElem::I32, 4),
        (TypedElem::U32, 4),
        (TypedElem::F32, 4),
        (TypedElem::F64, 8),
        (TypedElem::BigI64, 8),
        (TypedElem::BigU64, 8),
    ];
    for (variant, bpe) in expected {
        assert_eq!(variant.bytes_per_element(), *bpe,
            "{} must have BYTES_PER_ELEMENT = {}", variant.module(), bpe);
    }
}

#[test]
fn typed_array_variants_list_covers_all_11() {
    assert_eq!(js_typedarray_builtins::VARIANTS.len(), 11,
        "ECMA-262 §23.2 defines exactly 11 typed-array variants");
}

// ──────────────────────────────────────────────────────────────────────
// Contract 7: core-method surface coverage per MDN
// ──────────────────────────────────────────────────────────────────────

#[test]
fn array_surface_covers_mdn_prototype_methods() {
    // ECMA-262 §23.1.3 canonical prototype method set. If any method
    // is missing from IMPORTS, MDN examples for it will fail when
    // Phase D JS migration emits the call.
    let required: &[&str] = &[
        // Mutating
        "push", "pop", "shift", "unshift", "splice", "reverse", "sort",
        "fill", "copyWithin",
        // Non-mutating
        "slice", "concat", "indexOf", "lastIndexOf", "includes",
        "find", "findIndex", "findLast", "findLastIndex",
        "join", "flat", "flatMap",
        // Iteration
        "forEach", "map", "filter", "reduce", "reduceRight",
        "some", "every", "keys", "values", "entries",
        // ES2023 non-mutating variants
        "toReversed", "toSorted", "toSpliced", "with",
        // Access
        "at", "length",
    ];
    for name in required {
        assert!(js_array_builtins::IMPORTS.contains(name),
            "wasm:js-array missing required `{}` per ECMA-262 §23.1.3", name);
    }
}

#[test]
fn map_surface_covers_mdn_prototype_methods() {
    let required: &[&str] = &[
        "get", "set", "has", "delete", "clear", "size",
        "keys", "values", "entries", "forEach",
    ];
    for name in required {
        assert!(js_map_builtins::IMPORTS.contains(name),
            "wasm:js-map missing required `{}` per ECMA-262 §24.1.3", name);
    }
}

#[test]
fn set_surface_covers_mdn_prototype_methods_including_es2025() {
    let required: &[&str] = &[
        // Core
        "add", "has", "delete", "clear", "size",
        "values", "keys", "entries", "forEach",
        // ES2025 algebra
        "union", "intersection", "difference", "symmetricDifference",
        "isSubsetOf", "isSupersetOf", "isDisjointFrom",
    ];
    for name in required {
        assert!(js_set_builtins::IMPORTS.contains(name),
            "wasm:js-set missing required `{}` per ECMA-262 §24.2 + ES2025 set algebra", name);
    }
}

#[test]
fn dataview_has_all_get_and_set_variants() {
    // ECMA-262 §25.3.3 — every numeric element type is accessible
    // via both a getter and a setter.
    let element_types = &[
        "Int8", "Uint8", "Int16", "Uint16", "Int32", "Uint32",
        "BigInt64", "BigUint64", "Float32", "Float64",
    ];
    for ty in element_types {
        let getter = format!("get{}", ty);
        let setter = format!("set{}", ty);
        assert!(js_arraybuffer_builtins::DATAVIEW_IMPORTS
            .iter().any(|s| *s == getter.as_str()),
            "wasm:js-dataview missing {}", getter);
        assert!(js_arraybuffer_builtins::DATAVIEW_IMPORTS
            .iter().any(|s| *s == setter.as_str()),
            "wasm:js-dataview missing {}", setter);
    }
}

#[test]
fn object_surface_covers_mdn_static_and_prototype_methods() {
    let required: &[&str] = &[
        // Construction
        "new", "create", "fromEntries", "assign",
        // Access
        "get", "set", "has", "hasOwn", "delete",
        // Enumeration
        "keys", "values", "entries", "getOwnPropertyNames",
        // Descriptors
        "defineProperty", "defineProperties",
        "getOwnPropertyDescriptor", "getOwnPropertyDescriptors",
        // Prototype
        "getPrototypeOf", "setPrototypeOf",
        // Locking
        "freeze", "isFrozen", "seal", "isSealed",
        "preventExtensions", "isExtensible",
        // Comparison
        "is",
    ];
    for name in required {
        assert!(js_object_builtins::IMPORTS.contains(name),
            "wasm:js-object missing required `{}` per ECMA-262 §20.1", name);
    }
}

// ──────────────────────────────────────────────────────────────────────
// Contract 8: baseline size sanity
// ──────────────────────────────────────────────────────────────────────

#[test]
fn total_imported_surface_is_reasonable() {
    // Rough gauge: if the numbers drop precipitously something got
    // silently removed; if they balloon without a commit we should
    // notice.
    let array_n = js_array_builtins::IMPORTS.len();
    let object_n = js_object_builtins::IMPORTS.len();
    let map_n = js_map_builtins::IMPORTS.len();
    let set_n = js_set_builtins::IMPORTS.len();
    let weakmap_n = js_weakmap_builtins::WEAKMAP_IMPORTS.len();
    let weakset_n = js_weakmap_builtins::WEAKSET_IMPORTS.len();
    let ab_n = js_arraybuffer_builtins::ARRAYBUFFER_IMPORTS.len();
    let sab_n = js_arraybuffer_builtins::SHAREDARRAYBUFFER_IMPORTS.len();
    let dv_n = js_arraybuffer_builtins::DATAVIEW_IMPORTS.len();
    let ta_methods = js_typedarray_builtins::IMPORTS.len();

    assert!(array_n >= 40, "Array surface should be ~40+ methods, got {}", array_n);
    assert!(object_n >= 25, "Object surface should be 25+ methods, got {}", object_n);
    assert!(map_n >= 10, "Map surface should be 10+ methods, got {}", map_n);
    assert!(set_n >= 15, "Set surface (with ES2025) should be 15+, got {}", set_n);
    assert!(weakmap_n >= 5, "WeakMap surface should be 5+, got {}", weakmap_n);
    assert!(weakset_n >= 4, "WeakSet surface should be 4+, got {}", weakset_n);
    assert!(ab_n >= 10, "ArrayBuffer surface should be 10+, got {}", ab_n);
    assert!(sab_n >= 6, "SharedArrayBuffer surface should be 6+, got {}", sab_n);
    assert!(dv_n >= 22, "DataView surface should be 22+ (11 get + 11 set + helpers), got {}", dv_n);
    assert!(ta_methods >= 30, "TypedArray method set should be 30+, got {}", ta_methods);

    let total = array_n + object_n + map_n + set_n + weakmap_n + weakset_n
        + ab_n + sab_n + dv_n + ta_methods * 11;
    assert!(total >= 300,
        "Full js-builtin surface should be 300+ distinct (module, method) pairs, got {}",
        total);
}
