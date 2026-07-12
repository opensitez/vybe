//! Reflect.get/set with receiver — accessors, prototypes, non-object targets.

crate::js_cases! {
    reflect_get_on_data_property => {
        r#"const o={x:1}; console.log(Reflect.get(o,"x"));"#,
        ["1"]
    };

    reflect_get_invokes_getter_with_receiver => {
        r#"const o={get x(){return this._v;}, _v:9}; console.log(Reflect.get(o,"x",{_v:3}));"#,
        ["3"]
    };

    reflect_set_invokes_setter_with_receiver => {
        r#"const target={}; const recv={v:0}; const o={set x(val){this.v=val;}}; Object.setPrototypeOf(o,target); Reflect.set(o,"x",5,recv); console.log(recv.v);"#,
        ["5"]
    };

    reflect_get_finds_inherited_property => {
        r#"const p={a:1}; const o=Object.create(p); console.log(Reflect.get(o,"a"));"#,
        ["1"]
    };

    reflect_set_on_non_writable_data_returns_false => {
        r#"const o={}; Object.defineProperty(o,"x",{value:1,writable:false}); console.log(Reflect.set(o,"x",2));console.log(o.x);"#,
        ["false", "1"]
    };

    reflect_set_on_accessor_without_setter_returns_false => {
        r#"const o={get x(){return 1;}}; console.log(Reflect.set(o,"x",2));"#,
        ["false"]
    };

    reflect_get_symbol_key => {
        r#"const s=Symbol("k"); const o={[s]:7}; console.log(Reflect.get(o,s));"#,
        ["7"]
    };

    reflect_set_symbol_key => {
        r#"const s=Symbol("k"); const o={}; console.log(Reflect.set(o,s,8));console.log(o[s]);"#,
        ["true", "8"]
    };

    reflect_has_on_proxy_target => {
        r#"const o={a:1}; console.log(Reflect.has(o,"a"));console.log(Reflect.has(o,"b"));"#,
        ["true", "false"]
    };

    reflect_delete_non_configurable_returns_false => {
        r#"const o={}; Object.defineProperty(o,"x",{value:1,configurable:false}); console.log(Reflect.deleteProperty(o,"x"));"#,
        ["false"]
    };

    reflect_own_keys_includes_symbols => {
        r#"const s=Symbol("s"); const o={a:1,[s]:2}; const k=Reflect.ownKeys(o); console.log(k.includes("a"));console.log(k.includes(s));"#,
        ["true", "true"]
    };

    reflect_get_own_property_descriptor_returns_descriptor => {
        r#"const o={}; Object.defineProperty(o,"x",{value:1,enumerable:true}); const d=Reflect.getOwnPropertyDescriptor(o,"x"); console.log(d.value);console.log(d.enumerable);"#,
        ["1", "true"]
    };

    reflect_define_property_adds_new_key => {
        r#"const o={}; console.log(Reflect.defineProperty(o,"n",{value:2}));console.log(o.n);"#,
        ["true", "2"]
    };

    reflect_define_property_reject_invalid_descriptor => {
        r#"const o={}; Object.defineProperty(o,"x",{value:1,configurable:false}); console.log(Reflect.defineProperty(o,"x",{value:2,configurable:true}));"#,
        ["false"]
    };

    reflect_get_prototype_of_plain_object => {
        r#"console.log(Reflect.getPrototypeOf({})===Object.prototype);"#,
        ["true"]
    };

    reflect_set_prototype_of_changes_chain => {
        r#"const o={}; const p={tag:"p"}; Reflect.setPrototypeOf(o,p); console.log(Reflect.getPrototypeOf(o)===p);"#,
        ["true"]
    };

    reflect_is_extensible_true_on_new_object => {
        r#"console.log(Reflect.isExtensible({}));"#,
        ["true"]
    };

    reflect_prevent_extensions_blocks_add => {
        r#"const o={x:1}; Reflect.preventExtensions(o); console.log(Reflect.set(o,"y",2));console.log("y" in o);"#,
        ["false", "false"]
    };

    reflect_apply_with_array_like_args => {
        r#"function sum(a,b){return a+b;} console.log(Reflect.apply(sum,null,{0:3,1:4,length:2}));"#,
        ["7"]
    };

    reflect_construct_with_new_target => {
        r#"function F(v){this.v=v;} const i=Reflect.construct(F,[5]); console.log(i instanceof F);console.log(i.v);"#,
        ["true", "5"]
    };

    // Node-verified: §28.1.2 — the explicit newTarget (A) is exactly what
    // `new.target` observes through the ctor chain; node prints "A".
    reflect_construct_passes_new_target_to_operators => {
        r#"class A{constructor(){this.kind=new.target.name;}} class B extends A{} const i=Reflect.construct(B,[],A); console.log(i.kind);"#,
        ["A"]
    };

    // Node-verified: §28.1.5 step 1 — primitives are not Objects; no
    // wrapper coercion happens, it throws TypeError.
    reflect_get_on_primitive_string_wrapper_coercion => {
        r#"try{Reflect.get("hi","length");}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    // Node-verified: §28.1.12 step 1 — non-object target throws TypeError.
    reflect_set_on_primitive_returns_false => {
        r#"try{Reflect.set("hi","length",9);}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    reflect_get_on_null_throws => {
        r#"try{Reflect.get(null,"x");}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    reflect_set_on_null_throws => {
        r#"try{Reflect.set(null,"x",1);}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    reflect_get_accessor_receiver_changes_this_in_getter => {
        r#"const o={get val(){return this.tag;}}; console.log(Reflect.get(o,"val",{tag:"recv"}));"#,
        ["recv"]
    };

    reflect_set_accessor_receiver_changes_this_in_setter => {
        r#"const store={}; const o={set val(v){this.stored=v;}}; Reflect.set(o,"val",42,{stored:undefined,get stored(){return store.v},set stored(x){store.v=x;}}); console.log(store.v);"#,
        ["42"]
    };

    reflect_has_on_proxy_via_handler => {
        r#"const o=new Proxy({},{has(){return true;}}); console.log(Reflect.has(o,"missing"));"#,
        ["true"]
    };

    reflect_get_on_proxy_forwards_to_target => {
        r#"const o=new Proxy({x:1},{}); console.log(Reflect.get(o,"x"));"#,
        ["1"]
    };

    reflect_set_on_proxy_forwards_to_target => {
        r#"const t={}; const o=new Proxy(t,{}); Reflect.set(o,"a",1); console.log(t.a);"#,
        ["1"]
    };

    // Node-verified: ownKeys returns STRING keys ("0","1","length") —
    // includes(0) with a number is false (§28.1.10 / §10.4.2.4).
    reflect_own_keys_on_array_includes_length => {
        r#"const k=Reflect.ownKeys([1,2]); console.log(k.includes("length"));console.log(k.includes("0"));"#,
        ["true", "true"]
    };

    reflect_get_prototype_of_array => {
        r#"console.log(Reflect.getPrototypeOf([])===Array.prototype);"#,
        ["true"]
    };

    // Node-verified: §7.3.15 SetIntegrityLevel calls [[PreventExtensions]],
    // so a sealed object is NOT extensible.
    reflect_is_extensible_on_sealed_object => {
        r#"const o=Object.seal({}); console.log(Reflect.isExtensible(o));"#,
        ["false"]
    };

    reflect_prevent_extensions_on_frozen_object_still_false_extensible => {
        r#"const o=Object.freeze({}); console.log(Reflect.isExtensible(o));"#,
        ["false"]
    };

    reflect_delete_property_on_proxy_target => {
        r#"const t={a:1}; const o=new Proxy(t,{}); console.log(Reflect.deleteProperty(o,"a"));console.log("a" in t);"#,
        ["true", "false"]
    };

    reflect_define_property_on_non_object_throws => {
        r#"try{Reflect.defineProperty(1,"x",{value:1});}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    reflect_get_own_property_descriptor_missing_returns_undefined => {
        r#"console.log(Reflect.getOwnPropertyDescriptor({},"n")===undefined);"#,
        ["true"]
    };

    reflect_apply_throws_when_target_not_callable => {
        r#"try{Reflect.apply({},null,[]);}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    reflect_construct_throws_when_target_not_constructor => {
        r#"try{Reflect.construct(()=>{},[]);}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    reflect_get_integer_indexed_on_typed_array => {
        r#"const a=new Uint8Array([5]); console.log(Reflect.get(a,0));"#,
        ["5"]
    };

    reflect_set_integer_indexed_on_typed_array => {
        r#"const a=new Uint8Array(1); console.log(Reflect.set(a,0,9));console.log(a[0]);"#,
        ["true", "9"]
    };

    reflect_has_array_index => {
        r#"const a=[10]; console.log(Reflect.has(a,0));console.log(Reflect.has(a,"0"));"#,
        ["true", "true"]
    };

    // Node-verified: Date instances have NO own keys — internal slots are
    // not properties (§10.1.11); Reflect.ownKeys(new Date(0)) is [].
    reflect_own_keys_on_date_includes_internal_slots_as_keys => {
        r#"const k=Reflect.ownKeys(new Date(0)); console.log(k.length>0);"#,
        ["false"]
    };

    reflect_get_prototype_of_function => {
        r#"console.log(Reflect.getPrototypeOf(function(){})===Function.prototype);"#,
        ["true"]
    };

    reflect_set_prototype_of_function_allowed => {
        r#"const f=function(){}; const p={}; Reflect.setPrototypeOf(f,p); console.log(Reflect.getPrototypeOf(f)===p);"#,
        ["true"]
    };

    reflect_get_on_module_namespace_exotic => {
        r#"const o={}; Object.defineProperty(o,"locked",{value:1,configurable:false,enumerable:true}); console.log(Reflect.get(o,"locked"));"#,
        ["1"]
    };

    reflect_set_existing_property_on_non_extensible_succeeds => {
        r#"const o={x:1}; Reflect.preventExtensions(o); console.log(Reflect.set(o,"x",2));console.log(o.x);"#,
        ["true", "2"]
    };

    reflect_delete_on_non_configurable_returns_false => {
        r#"const o=Object.freeze({x:1}); console.log(Reflect.deleteProperty(o,"x"));"#,
        ["false"]
    };

    reflect_define_property_on_extensible_object_adds_symbol => {
        r#"const s=Symbol("d"); const o={}; Reflect.defineProperty(o,s,{value:1}); console.log(Reflect.get(o,s));"#,
        ["1"]
    };

    reflect_get_with_receiver_on_inherited_accessor => {
        r#"const base={_v:1,get g(){return this._v;}}; const o=Object.create(base); console.log(Reflect.get(o,"g",{_v:100}));"#,
        ["100"]
    };
}
