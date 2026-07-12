//! Object integrity levels — preventExtensions, seal, freeze error behavior.

crate::js_cases! {
    prevent_extensions_blocks_new_property => {
        r#"const o={a:1}; Object.preventExtensions(o); o.b=2; console.log("b" in o);"#,
        ["false"]
    };

    prevent_extensions_allows_existing_write => {
        r#"const o={a:1}; Object.preventExtensions(o); o.a=9; console.log(o.a);"#,
        ["9"]
    };

    prevent_extensions_allows_delete_existing => {
        r#"const o={a:1}; Object.preventExtensions(o); delete o.a; console.log("a" in o);"#,
        ["false"]
    };

    seal_blocks_add_and_delete => {
        r#"const o={a:1}; Object.seal(o); o.b=2; delete o.a; console.log("a" in o);console.log("b" in o);"#,
        ["true", "false"]
    };

    seal_allows_modify_existing_data => {
        r#"const o={a:1}; Object.seal(o); o.a=5; console.log(o.a);"#,
        ["5"]
    };

    freeze_blocks_add_delete_and_write => {
        r#"const o={a:1}; Object.freeze(o); o.a=2; o.b=3; delete o.a; console.log(o.a);"#,
        ["1"]
    };

    freeze_returns_same_object => {
        r#"const o={}; console.log(Object.freeze(o)===o);"#,
        ["true"]
    };

    is_extensible_false_after_prevent_extensions => {
        r#"const o={}; Object.preventExtensions(o); console.log(Object.isExtensible(o));"#,
        ["false"]
    };

    is_sealed_true_after_seal => {
        r#"console.log(Object.isSealed(Object.seal({})));"#,
        ["true"]
    };

    is_frozen_true_after_freeze => {
        r#"console.log(Object.isFrozen(Object.freeze({})));"#,
        ["true"]
    };

    freeze_array_blocks_push => {
        r#"const a=Object.freeze([1,2]); try{a.push(3); console.log("ok");}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    seal_array_blocks_push_allows_index_write => {
        r#"const a=Object.seal([1,2]); try{a.push(3);}catch(e){} a[0]=9; console.log(a[0]);"#,
        ["9"]
    };

    prevent_extensions_on_empty_object => {
        r#"const o=Object.preventExtensions({}); console.log(Object.isExtensible(o));"#,
        ["false"]
    };

    // Node-verified: Object.defineProperty THROWS TypeError on failure
    // (§20.1.2.4); it never returns false — that's Reflect.defineProperty.
    define_property_on_non_extensible_fails_for_new_key => {
        r#"const o=Object.preventExtensions({}); try{Object.defineProperty(o,"x",{value:1,configurable:true,enumerable:true,writable:true});}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    define_property_on_sealed_fails_for_new_key => {
        r#"const o=Object.seal({}); try{Object.defineProperty(o,"n",{value:1,configurable:true,enumerable:true,writable:true});}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    delete_on_sealed_property_fails => {
        r#"const o=Object.seal({x:1}); console.log(delete o.x);console.log(o.x);"#,
        ["false", "1"]
    };

    write_on_frozen_property_fails_silently_in_sloppy => {
        r#"const o=Object.freeze({x:1}); o.x=2; console.log(o.x);"#,
        ["1"]
    };

    strict_write_on_frozen_throws_in_strict => {
        r#""use strict"; const o=Object.freeze({x:1}); try{o.x=2;}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    freeze_nested_object_not_deep_by_default => {
        r#"const o=Object.freeze({n:{v:1}}); o.n.v=2; console.log(o.n.v);"#,
        ["2"]
    };

    prevent_extensions_then_seal_still_sealed => {
        r#"const o=Object.preventExtensions({a:1}); Object.seal(o); console.log(Object.isSealed(o));"#,
        ["true"]
    };

    seal_then_freeze_is_frozen => {
        r#"const o=Object.freeze(Object.seal({a:1})); console.log(Object.isFrozen(o));"#,
        ["true"]
    };

    reflect_prevent_extensions_matches_object => {
        r#"const o={}; Reflect.preventExtensions(o); console.log(Object.isExtensible(o));"#,
        ["false"]
    };

    // Node-verified: Object.assign uses [[Set]] with throw semantics —
    // a new key on a non-extensible target throws TypeError (§20.1.2.1),
    // it does not skip silently.
    object_assign_to_non_extensible_skips_new_keys => {
        r#"const o=Object.preventExtensions({}); try{Object.assign(o,{a:1});}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    object_assign_to_sealed_skips_new_keys => {
        r#"const o=Object.seal({}); try{Object.assign(o,{b:2});}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    json_stringify_on_frozen_object => {
        r#"console.log(JSON.stringify(Object.freeze({a:1})));"#,
        ["{\"a\":1}"]
    };

    keys_on_sealed_returns_existing => {
        r#"console.log(Object.keys(Object.seal({x:1,y:2})).sort().join(","));"#,
        ["x,y"]
    };

    get_own_property_descriptor_on_frozen => {
        r#"const d=Object.getOwnPropertyDescriptor(Object.freeze({x:1}),"x"); console.log(d.writable);"#,
        ["false"]
    };

    // Node-verified: since ES2015, seal/freeze on a non-object is a
    // NO-OP returning the argument (§20.1.2.20/§20.1.2.7 step 1) — the
    // ES5 TypeError is gone.
    seal_non_object_throws => {
        r#"console.log(Object.seal(1));"#,
        ["1"]
    };

    freeze_null_throws => {
        r#"console.log(Object.freeze(null));"#,
        ["null"]
    };

    prevent_extensions_on_function_allowed => {
        r#"const f=function(){}; Object.preventExtensions(f); console.log(Object.isExtensible(f));"#,
        ["false"]
    };

    freeze_date_instance => {
        r#"const d=Object.freeze(new Date(0)); console.log(Object.isFrozen(d));"#,
        ["true"]
    };

    seal_with_symbol_key_property => {
        r#"const s=Symbol("k"); const o=Object.seal({[s]:1}); o[s]=2; console.log(o[s]);"#,
        ["2"]
    };

    freeze_with_non_configurable_existing => {
        r#"const o={}; Object.defineProperty(o,"x",{value:1,configurable:false,writable:true}); Object.freeze(o); console.log(Object.isFrozen(o));"#,
        ["true"]
    };

    prevent_extensions_idempotent => {
        r#"const o={}; Object.preventExtensions(o); Object.preventExtensions(o); console.log(Object.isExtensible(o));"#,
        ["false"]
    };

    is_sealed_false_on_new_object => {
        r#"console.log(Object.isSealed({}));"#,
        ["false"]
    };

    is_frozen_false_on_new_object => {
        r#"console.log(Object.isFrozen({}));"#,
        ["false"]
    };

    // Node-verified: Object.setPrototypeOf on a non-extensible target
    // THROWS TypeError (§20.1.2.22); false is Reflect.setPrototypeOf.
    set_prototype_on_non_extensible_fails => {
        r#"const o=Object.preventExtensions({}); try{Object.setPrototypeOf(o,{});}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    array_buffer_not_freezable_as_object => {
        r#"const buf=new ArrayBuffer(4); console.log(Object.isExtensible(buf));"#,
        ["true"]
    };

    // Node-verified: FREEZING a typed array with elements throws
    // TypeError immediately ("Cannot freeze array buffer views with
    // elements") — integer-indexed elements can't be made non-writable.
    typed_array_freeze_blocks_length_change_via_methods => {
        r#"try{Object.freeze(new Uint8Array([1]));}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };
}
