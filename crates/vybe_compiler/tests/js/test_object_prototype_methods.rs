//! Builtin prototype method coverage — distinct behaviors only.
crate::js_cases! {
    object_tostring_tag_undefined => {
        r#"console.log(Object.prototype.toString.call(undefined).slice(8,-1));"#,
        ["Undefined"]
    };

    object_tostring_tag_null => {
        r#"console.log(Object.prototype.toString.call(null).slice(8,-1));"#,
        ["Null"]
    };

    object_tostring_tag_true => {
        r#"console.log(Object.prototype.toString.call(true).slice(8,-1));"#,
        ["Boolean"]
    };

    object_tostring_tag_false => {
        r#"console.log(Object.prototype.toString.call(false).slice(8,-1));"#,
        ["Boolean"]
    };

    object_tostring_tag_zero => {
        r#"console.log(Object.prototype.toString.call(0).slice(8,-1));"#,
        ["Number"]
    };

    object_tostring_tag_nan => {
        r#"console.log(Object.prototype.toString.call(NaN).slice(8,-1));"#,
        ["Number"]
    };

    object_tostring_tag_infinity => {
        r#"console.log(Object.prototype.toString.call(Infinity).slice(8,-1));"#,
        ["Number"]
    };

    object_tostring_tag_string => {
        r#"console.log(Object.prototype.toString.call("hello").slice(8,-1));"#,
        ["String"]
    };

    object_tostring_tag_array => {
        r#"console.log(Object.prototype.toString.call([1,2]).slice(8,-1));"#,
        ["Array"]
    };

    object_tostring_tag_plain_object => {
        r#"console.log(Object.prototype.toString.call(({})).slice(8,-1));"#,
        ["Object"]
    };

    object_tostring_tag_function => {
        r#"console.log(Object.prototype.toString.call(function f(){}).slice(8,-1));"#,
        ["Function"]
    };

    object_tostring_tag_date => {
        r#"console.log(Object.prototype.toString.call(new Date(0)).slice(8,-1));"#,
        ["Date"]
    };

    object_tostring_tag_regex => {
        r#"console.log(Object.prototype.toString.call(/abc/gi).slice(8,-1));"#,
        ["RegExp"]
    };

    object_tostring_tag_map => {
        r#"console.log(Object.prototype.toString.call(new Map()).slice(8,-1));"#,
        ["Map"]
    };

    object_tostring_tag_set => {
        r#"console.log(Object.prototype.toString.call(new Set()).slice(8,-1));"#,
        ["Set"]
    };

    object_tostring_tag_promise => {
        r#"console.log(Object.prototype.toString.call(Promise.resolve(1)).slice(8,-1));"#,
        ["Promise"]
    };

    object_tostring_tag_error => {
        r#"console.log(Object.prototype.toString.call(new Error('x')).slice(8,-1));"#,
        ["Error"]
    };

    object_tostring_tag_math => {
        r#"console.log(Object.prototype.toString.call(Math).slice(8,-1));"#,
        ["Object"]
    };

    object_tostring_tag_json => {
        r#"console.log(Object.prototype.toString.call(JSON).slice(8,-1));"#,
        ["Object"]
    };

    object_tostring_tag_weakmap => {
        r#"console.log(Object.prototype.toString.call(new WeakMap()).slice(8,-1));"#,
        ["WeakMap"]
    };

    object_tostring_tag_weakset => {
        r#"console.log(Object.prototype.toString.call(new WeakSet()).slice(8,-1));"#,
        ["WeakSet"]
    };

    object_tostring_tag_int8array => {
        r#"console.log(Object.prototype.toString.call(new Int8Array(1)).slice(8,-1));"#,
        ["Int8Array"]
    };

    object_tostring_tag_arraybuffer => {
        r#"console.log(Object.prototype.toString.call(new ArrayBuffer(8)).slice(8,-1));"#,
        ["ArrayBuffer"]
    };

    object_tostring_tag_dataview => {
        r#"console.log(Object.prototype.toString.call(new DataView(new ArrayBuffer(4))).slice(8,-1));"#,
        ["DataView"]
    };

    object_tostring_tag_bigint => {
        r#"console.log(Object.prototype.toString.call(1n).slice(8,-1));"#,
        ["BigInt"]
    };

    plain_object_valueof_returns_self => {
        r#"const o={a:1}; console.log(o.valueOf()===o);"#,
        ["true"]
    };

    array_valueof_returns_self => {
        r#"const a=[1]; console.log(a.valueOf()===a);"#,
        ["true"]
    };

    function_valueof_returns_self => {
        r#"function f(){} console.log(f.valueOf()===f);"#,
        ["true"]
    };

    date_valueof_returns_timestamp => {
        r#"const d=new Date(1000); console.log(d.valueOf());"#,
        ["1000"]
    };

    number_primitive_valueof => {
        r#"console.log((42).valueOf());"#,
        ["42"]
    };

    string_primitive_valueof => {
        r#"console.log("hi".valueOf());"#,
        ["hi"]
    };

    regex_valueof_returns_self => {
        r#"const r=/a/; console.log(r.valueOf()===r);"#,
        ["true"]
    };

    valueof_overridden_returns_custom => {
        r#"const o={valueOf(){return 99}}; console.log(o.valueOf()); console.log(+o);"#,
        ["99", "99"]
    };

    tostring_default_plain_object => {
        r#"console.log({}.toString());"#,
        ["[object Object]"]
    };

    tostring_array_joins_elements => {
        r#"console.log([1,2,3].toString());"#,
        ["1,2,3"]
    };

    tostring_function_contains_name => {
        r#"function foo(){} console.log(foo.toString().includes("foo"));"#,
        ["true"]
    };

    tostring_overridden_on_object => {
        r#"const o={toString(){return "custom"}}; console.log(String(o));"#,
        ["custom"]
    };

    tostring_call_arguments_object => {
        r#"function f(){console.log(Object.prototype.toString.call(arguments).slice(8,-1));} f(1);"#,
        ["Arguments"]
    };

    hasownproperty_symbol_own_key => {
        r#"const s=Symbol("k"); const o={[s]:1}; console.log(o.hasOwnProperty(s));"#,
        ["true"]
    };

    hasownproperty_symbol_inherited_false => {
        r#"const s=Symbol("k"); const p={[s]:1}; const o=Object.create(p); console.log(o.hasOwnProperty(s));"#,
        ["false"]
    };

    hasownproperty_array_index_own => {
        r#"const a=[10,20]; console.log(a.hasOwnProperty(0)); console.log(a.hasOwnProperty(1));"#,
        ["true", "true"]
    };

    hasownproperty_array_length => {
        r#"const a=[1]; console.log(a.hasOwnProperty("length"));"#,
        ["true"]
    };

    hasownproperty_inherited_tostring_false => {
        r#"const o={}; console.log(o.hasOwnProperty("toString"));"#,
        ["false"]
    };

    hasownproperty_null_proto_via_call => {
        r#"const o=Object.create(null); o.x=1; console.log(Object.prototype.hasOwnProperty.call(o,"x")); console.log(Object.prototype.hasOwnProperty.call(o,"toString"));"#,
        ["true", "false"]
    };

    hasownproperty_after_instance_override => {
        r#"const o={hasOwnProperty(){return false}}; console.log(o.hasOwnProperty("x")); console.log(Object.prototype.hasOwnProperty.call(o,"hasOwnProperty"));"#,
        ["false", "true"]
    };

    hasownproperty_non_configurable_own => {
        r#"const o={}; Object.defineProperty(o,"n",{value:1,configurable:false}); console.log(o.hasOwnProperty("n"));"#,
        ["true"]
    };

    hasownproperty_getter_descriptor => {
        r#"const o={}; Object.defineProperty(o,"g",{get(){return 1}}); console.log(o.hasOwnProperty("g"));"#,
        ["true"]
    };

    hasownproperty_prototype_method_not_own => {
        r#"function C(){} C.prototype.m=function(){}; const c=new C(); console.log(c.hasOwnProperty("m"));"#,
        ["false"]
    };

    hasownproperty_constructor_on_function => {
        r#"function F(){} console.log(F.hasOwnProperty("prototype"));"#,
        ["true"]
    };

    hasownproperty_numeric_string_key => {
        r#"const o={"0":1}; console.log(o.hasOwnProperty(0)); console.log(o.hasOwnProperty("0"));"#,
        ["true", "true"]
    };

    hasownproperty_after_delete => {
        r#"const o={a:1}; delete o.a; console.log(o.hasOwnProperty("a"));"#,
        ["false"]
    };

    // §24.1.3.10: Map's `size` is an ACCESSOR on Map.prototype, never an
    // own property of instances — node-verified false.
    hasownproperty_map_size_own => {
        r#"const m=new Map(); console.log(m.hasOwnProperty("size"));"#,
        ["false"]
    };

    isprototypeof_object_on_plain => {
        r#"console.log(Object.prototype.isPrototypeOf({}));"#,
        ["true"]
    };

    isprototypeof_object_on_null_proto => {
        r#"console.log(Object.prototype.isPrototypeOf(Object.create(null)));"#,
        ["false"]
    };

    isprototypeof_array_on_array => {
        r#"console.log(Array.prototype.isPrototypeOf([]));"#,
        ["true"]
    };

    isprototypeof_array_on_object_false => {
        r#"console.log(Array.prototype.isPrototypeOf({}));"#,
        ["false"]
    };

    isprototypeof_function_on_arrow => {
        r#"console.log(Function.prototype.isPrototypeOf(()=>{}));"#,
        ["true"]
    };

    isprototypeof_function_on_object_false => {
        r#"console.log(Function.prototype.isPrototypeOf({}));"#,
        ["false"]
    };

    isprototypeof_direct_parent => {
        r#"const p={}; const c=Object.create(p); console.log(p.isPrototypeOf(c)); console.log({}.isPrototypeOf(c));"#,
        ["true", "false"]
    };

    isprototypeof_after_setprototypeof => {
        r#"const a={}; const b={}; const c={}; Object.setPrototypeOf(c,a); console.log(a.isPrototypeOf(c)); Object.setPrototypeOf(c,b); console.log(a.isPrototypeOf(c)); console.log(b.isPrototypeOf(c));"#,
        ["true", "false", "true"]
    };

    isprototypeof_boxed_number => {
        r#"console.log(Number.prototype.isPrototypeOf(Object(7)));"#,
        ["true"]
    };

    isprototypeof_boxed_string => {
        r#"console.log(String.prototype.isPrototypeOf(Object("x")));"#,
        ["true"]
    };

    isprototypeof_boxed_boolean => {
        r#"console.log(Boolean.prototype.isPrototypeOf(Object(false)));"#,
        ["true"]
    };

    isprototypeof_date_instance => {
        r#"console.log(Date.prototype.isPrototypeOf(new Date()));"#,
        ["true"]
    };

    isprototypeof_regexp_instance => {
        r#"console.log(RegExp.prototype.isPrototypeOf(/a/));"#,
        ["true"]
    };

    propertyisenumerable_own_data => {
        r#"const o={a:1}; console.log(o.propertyIsEnumerable("a"));"#,
        ["true"]
    };

    propertyisenumerable_inherited_false => {
        r#"const p={a:1}; const o=Object.create(p); console.log(o.propertyIsEnumerable("a"));"#,
        ["false"]
    };

    propertyisenumerable_nonenumerable_own => {
        r#"const o={}; Object.defineProperty(o,"h",{value:1,enumerable:false}); console.log(o.propertyIsEnumerable("h"));"#,
        ["false"]
    };

    propertyisenumerable_symbol_enumerable => {
        r#"const s=Symbol("e"); const o={[s]:1}; console.log(o.propertyIsEnumerable(s));"#,
        ["true"]
    };

    propertyisenumerable_symbol_nonenumerable => {
        r#"const s=Symbol("h"); const o={}; Object.defineProperty(o,s,{value:1,enumerable:false}); console.log(o.propertyIsEnumerable(s));"#,
        ["false"]
    };

    propertyisenumerable_array_index => {
        r#"const a=[1,2]; console.log(a.propertyIsEnumerable(0)); console.log(a.propertyIsEnumerable("length"));"#,
        ["true", "false"]
    };

    propertyisenumerable_function_prototype => {
        r#"function demo(){} console.log(demo.propertyIsEnumerable("prototype"));"#,
        ["false"]
    };

    propertyisenumerable_string_char_index => {
        r#"const s="ab"; console.log(s.propertyIsEnumerable(0)); console.log(s.propertyIsEnumerable("0"));"#,
        ["true", "true"]
    };

    propertyisenumerable_length_on_array => {
        r#"const a=[1,2,3]; console.log(a.propertyIsEnumerable("length"));"#,
        ["false"]
    };

    propertyisenumerable_after_override => {
        r#"const o={propertyIsEnumerable(){return true}}; console.log(o.propertyIsEnumerable("missing"));"#,
        ["true"]
    };

    proto_getter_matches_getprototypeof => {
        r#"const p={a:1}; const o=Object.create(p); console.log(o.__proto__===p); console.log(Object.getPrototypeOf(o)===p);"#,
        ["true", "true"]
    };

    proto_setter_changes_lookup => {
        r#"const o={}; const p={v:1}; o.__proto__=p; console.log(o.v);"#,
        ["1"]
    };

    proto_setter_to_null => {
        r#"const o={x:1}; o.__proto__=null; console.log(o.x); console.log(Object.getPrototypeOf(o));"#,
        ["1", "null"]
    };

    proto_literal_in_object_literal => {
        r#"const p={m(){return 9}}; const o={__proto__:p}; console.log(o.m());"#,
        ["9"]
    };

    proto_delete_restores_original => {
        r#"const o={}; const orig=Object.getPrototypeOf(o); o.__proto__={z:1}; delete o.__proto__; console.log(Object.getPrototypeOf(o)===orig);"#,
        ["true"]
    };

    proto_get_on_function => {
        r#"function f(){} console.log(f.__proto__===Function.prototype);"#,
        ["true"]
    };

    proto_set_shadows_inherited => {
        r#"const p={x:1}; const o=Object.create(p); o.__proto__={x:2}; console.log(o.x);"#,
        ["2"]
    };

}
