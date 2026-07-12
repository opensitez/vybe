//! Proxy invariant violations and trap error paths — non-configurable, non-extensible.

crate::js_cases! {
    proxy_get_returns_value_for_target_property => {
        r#"const p=new Proxy({x:1},{}); console.log(p.x);"#,
        ["1"]
    };

    proxy_set_forwards_to_target => {
        r#"const t={}; const p=new Proxy(t,{}); p.a=2; console.log(t.a);"#,
        ["2"]
    };

    proxy_get_trap_can_synthesize_property => {
        r#"const p=new Proxy({},{get(t,k){return k+"!";}}); console.log(p.hello);"#,
        ["hello!"]
    };

    proxy_set_trap_can_reject_assignment => {
        r#"const p=new Proxy({},{set(){return false;}}); p.x=1; console.log(p.x);"#,
        ["undefined"]
    };

    proxy_has_trap_hides_property => {
        r#"const p=new Proxy({a:1},{has(){return false;}}); console.log("a" in p);"#,
        ["false"]
    };

    proxy_delete_trap_can_block_delete => {
        r#"const t={a:1}; const p=new Proxy(t,{deleteProperty(){return false;}}); delete p.a; console.log(t.a);"#,
        ["1"]
    };

    proxy_apply_trap_on_function_target => {
        r#"const fn=function(a,b){return a+b;}; const p=new Proxy(fn,{apply(t,_,args){return args[0]*args[1];}}); console.log(p(3,4));"#,
        ["12"]
    };

    proxy_construct_trap_creates_instance => {
        r#"class C{constructor(v){this.v=v;}} const p=new Proxy(C,{construct(t,args){return new t(args[0]*2);}}); console.log(new p(3).v);"#,
        ["6"]
    };

    proxy_get_own_property_descriptor_forwards => {
        r#"const t={}; Object.defineProperty(t,"x",{value:1}); const d=Object.getOwnPropertyDescriptor(new Proxy(t,{}),"x"); console.log(d.value);"#,
        ["1"]
    };

    proxy_define_property_trap_adds_key => {
        r#"const t={}; const p=new Proxy(t,{defineProperty(t,k,d){Object.defineProperty(t,k,d); return true;}}); Object.defineProperty(p,"n",{value:5}); console.log(t.n);"#,
        ["5"]
    };

    proxy_own_keys_trap_can_filter => {
        r#"const p=new Proxy({a:1,b:2},{ownKeys(){return ["a"];}}); console.log(Reflect.ownKeys(p).join(","));"#,
        ["a"]
    };

    proxy_get_prototype_of_forwards => {
        r#"const proto={tag:1}; const t=Object.create(proto); console.log(Reflect.getPrototypeOf(new Proxy(t,{}))===proto);"#,
        ["true"]
    };

    proxy_set_prototype_of_forwards => {
        r#"const t={}; const np={}; const p=new Proxy(t,{}); Object.setPrototypeOf(p,np); console.log(Object.getPrototypeOf(t)===np);"#,
        ["true"]
    };

    proxy_is_extensible_forwards => {
        r#"const t={}; Object.preventExtensions(t); console.log(Object.isExtensible(new Proxy(t,{})));"#,
        ["false"]
    };

    proxy_prevent_extensions_forwards => {
        r#"const t={}; const p=new Proxy(t,{}); Object.preventExtensions(p); console.log(Object.isExtensible(t));"#,
        ["false"]
    };

    // Node-verified: §10.5.5 — reporting a DIFFERENT value for a
    // non-configurable own target property throws TypeError (the trap
    // result is not silently replaced by the target's descriptor).
    proxy_non_configurable_invariant_on_get_descriptor => {
        r#"const t={}; Object.defineProperty(t,"x",{value:1,configurable:false}); const p=new Proxy(t,{getOwnPropertyDescriptor(){return {value:2,configurable:false,enumerable:true,writable:true};}}); try{Object.getOwnPropertyDescriptor(p,"x");}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    proxy_revocable_stops_forwarding_after_revoke => {
        r#"const {proxy,revoke}=Proxy.revocable({x:1},{}); revoke(); try{console.log(proxy.x);}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    proxy_get_trap_throw_propagates => {
        r#"const p=new Proxy({},{get(){throw new Error("trap");}}); try{console.log(p.a);}catch(e){console.log(e.message);}"#,
        ["trap"]
    };

    proxy_set_trap_throw_propagates => {
        r#"const p=new Proxy({},{set(){throw new Error("set");}}); try{p.x=1;}catch(e){console.log(e.message);}"#,
        ["set"]
    };

    proxy_array_target_length_property => {
        r#"const p=new Proxy([1,2],{}); console.log(p.length);"#,
        ["2"]
    };

    proxy_function_target_callable => {
        r#"const p=new Proxy(function(){return 1;},{}); console.log(p());"#,
        ["1"]
    };

    proxy_in_operator_with_has_trap => {
        r#"const p=new Proxy({a:1},{has(t,k){return k==="a";}}); console.log("a" in p);console.log("b" in p);"#,
        ["true", "false"]
    };

    proxy_get_for_symbol_key => {
        r#"const s=Symbol("k"); const p=new Proxy({[s]:9},{}); console.log(p[s]);"#,
        ["9"]
    };

    proxy_set_return_false_strict_throws => {
        r#""use strict"; const p=new Proxy({},{set(){return false;}}); try{p.x=1;}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    proxy_construct_without_new_on_proxy => {
        r#"class C{} const p=new Proxy(C,{}); console.log(new p instanceof C);"#,
        ["true"]
    };

    proxy_get_trap_uses_receiver => {
        r#"const p=new Proxy({get x(){return this===p;}},{get(t,k,r){return k==="x"?r===p:false;}}); console.log(p.x);"#,
        ["true"]
    };

    proxy_target_can_be_null_prototype_object => {
        r#"const t=Object.create(null); t.a=1; console.log(new Proxy(t,{}).a);"#,
        ["1"]
    };

    proxy_define_property_non_configurable_invariant => {
        r#"const t={}; Object.defineProperty(t,"x",{value:1,configurable:false}); const p=new Proxy(t,{}); console.log(Object.getOwnPropertyDescriptor(p,"x").configurable);"#,
        ["false"]
    };

    proxy_get_own_property_descriptor_undefined_for_missing => {
        r#"console.log(Object.getOwnPropertyDescriptor(new Proxy({},{}),"n"));"#,
        ["undefined"]
    };

    proxy_apply_with_this_argument => {
        r#"function f(){return this.v;} const p=new Proxy(f,{apply(t,recv){return recv.v;}}); console.log(p.call({v:8}));"#,
        ["8"]
    };

    proxy_construct_return_non_object_throws => {
        r#"const p=new Proxy(function(){},{construct(){return 1;}}); try{new p();}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    proxy_is_array_on_array_target => {
        r#"console.log(Array.isArray(new Proxy([],{})));"#,
        ["true"]
    };

    proxy_json_stringify_uses_target => {
        r#"console.log(JSON.stringify(new Proxy({a:1},{})));"#,
        ["{\"a\":1}"]
    };

    proxy_delete_non_configurable_returns_false => {
        r#"const t={}; Object.defineProperty(t,"x",{value:1,configurable:false}); console.log(Reflect.deleteProperty(new Proxy(t,{}),"x"));"#,
        ["false"]
    };

    proxy_own_keys_includes_non_enumerable => {
        r#"const t={}; Object.defineProperty(t,"h",{value:1,enumerable:false}); const k=Reflect.ownKeys(new Proxy(t,{})); console.log(k.includes("h"));"#,
        ["true"]
    };

    proxy_get_trap_for_string_target_coercion => {
        r#"const p=new Proxy({},{get(t,k){return String(k).length;}}); console.log(p.hello);"#,
        ["5"]
    };

    proxy_set_on_frozen_target_fails => {
        r#"const t=Object.freeze({}); const p=new Proxy(t,{}); console.log(Reflect.set(p,"x",1));"#,
        ["false"]
    };

    proxy_has_trap_return_non_boolean_coerces => {
        r#"const p=new Proxy({},{has(){return 1;}}); console.log("any" in p);"#,
        ["true"]
    };

    proxy_get_prototype_of_on_function => {
        r#"console.log(Reflect.getPrototypeOf(new Proxy(function(){},{}))===Function.prototype);"#,
        ["true"]
    };

    proxy_two_level_nested_get => {
        r#"const inner=new Proxy({v:1},{}); const outer=new Proxy({inner},{get(t,k){return t[k];}}); console.log(outer.inner.v);"#,
        ["1"]
    };
}
