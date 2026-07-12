// Arrow functions (§15.3) — lexical this, no [[Construct]], no own
// prototype/arguments, metadata, and parameter forms. Every expectation
// node-verified.
crate::js_cases! {
    arrow_lexical_this_in_method => {
        r#"const obj={v:42,get(){const a=()=>this.v;return a();}}; console.log(obj.get());"#,
        ["42"]
    };

    arrow_extracted_keeps_lexical_this => {
        r#"const holder={v:7,make(){return ()=>this.v;}}; const ext=holder.make(); console.log(ext());"#,
        ["7"]
    };

    arrow_call_cannot_rebind_this => {
        r#"const arrow=()=>this; console.log(arrow.call({x:9})===arrow.call({x:1}));"#,
        ["true"]
    };

    arrow_bind_binds_args_not_this => {
        r#"const add=(a,b)=>a+b; const inc=add.bind({ignored:1},1); console.log(inc(5));"#,
        ["6"]
    };

    block_bodied_arrow_ignores_this_arg => {
        r#"const blk=function(){ return (() => { return this; }); }.call({tag:"outer"})(); console.log(blk && blk.tag);"#,
        ["outer"]
    };

    arrow_has_no_own_prototype_property => {
        r#"console.log((()=>{}).hasOwnProperty("prototype"));"#,
        ["false"]
    };

    arrow_new_throws_type_error => {
        r#"try{ new (()=>{})(); console.log("constructed"); }catch(e){ console.log(e instanceof TypeError); }"#,
        ["true"]
    };

    arrow_typeof_is_function => {
        r#"console.log(typeof (()=>{}));"#,
        ["function"]
    };

    arrow_instanceof_function => {
        r#"console.log((()=>{}) instanceof Function);"#,
        ["true"]
    };

    arrow_name_inferred_from_variable => {
        r#"const jump=()=>{}; console.log(jump.name);"#,
        ["jump"]
    };

    arrow_length_counts_plain_params => {
        r#"console.log(((a,b,c)=>0).length);"#,
        ["3"]
    };

    arrow_length_zero_at_leading_default => {
        r#"console.log(((a=1,b)=>0).length);"#,
        ["0"]
    };

    arrow_length_zero_for_rest_only => {
        r#"console.log(((...r)=>0).length);"#,
        ["0"]
    };

    arrow_parenthesized_object_literal_body => {
        r#"const mk=(x)=>({k:x}); console.log(mk(3).k);"#,
        ["3"]
    };

    arrow_default_parameter => {
        r#"const d=(a=5)=>a; console.log(d()); console.log(d(9));"#,
        ["5", "9"]
    };

    arrow_rest_parameter => {
        r#"const r=(...xs)=>xs.length; console.log(r(1,2,3));"#,
        ["3"]
    };

    arrow_destructured_parameters => {
        r#"const ds=({a},[b])=>a+b; console.log(ds({a:1},[2]));"#,
        ["3"]
    };

    nested_arrows_share_lexical_this => {
        r#"const deep={v:3,run(){return (()=>(()=>this.v)())();}}; console.log(deep.run());"#,
        ["3"]
    };

    async_arrow_awaits => {
        r#"(async()=>{ const f=async(x)=>x*2; console.log(await f(21)); })();"#,
        ["42"]
    };

    class_field_arrow_binds_instance => {
        r#"class C{v=5;get=()=>this.v;} const c=new C(); const g=c.get; console.log(g());"#,
        ["5"]
    };

    arrow_iife => {
        r#"console.log(((x)=>x+1)(41));"#,
        ["42"]
    };

    arrow_to_string_contains_arrow => {
        r#"console.log(Function.prototype.toString.call(()=>1).includes("=>"));"#,
        ["true"]
    };

    arrow_sees_enclosing_arguments => {
        r#"function outerFn(){ const a=()=>arguments[0]; return a(); } console.log(outerFn(99));"#,
        ["99"]
    };

    // §15.3: arrows have LEXICAL new.target — inside a construction it is
    // the enclosing constructor, in a plain call undefined.
    arrow_lexical_new_target_in_ctor => {
        r#"function NT(){ this.t=(()=>new.target)(); } console.log(new NT().t === NT);"#,
        ["true"]
    };

    arrow_lexical_new_target_undefined_in_plain_call => {
        r#"console.log((function(){ return (()=>new.target)(); })() === undefined);"#,
        ["true"]
    };

    arrow_lexical_super_in_method => {
        r#"class B{who(){return "base";}} class D extends B{who(){const a=()=>super.who();return a()+"+d";}} console.log(new D().who());"#,
        ["base+d"]
    };

    single_param_arrow_without_parens => {
        r#"const sp=x=>x*3; console.log(sp(4));"#,
        ["12"]
    };

    async_single_param_arrow_without_parens => {
        r#"(async()=>{ const f=async x=>x+1; console.log(await f(1)); })();"#,
        ["2"]
    };

    // Arrow as object property: `this` is the OUTER scope, never the
    // holder object.
    arrow_property_this_is_outer_scope => {
        r#"const o={v:1,f:()=>this}; console.log(o.f()===o.f()); console.log(o.f()!==o);"#,
        ["true", "true"]
    };

    arrow_trailing_comma_in_params => {
        r#"const tc=(a,b,)=>a+b; console.log(tc(1,2));"#,
        ["3"]
    };

    generator_arrow_is_syntax_error => {
        r#"try{ eval("const g = *() => 1;"); console.log("ok"); }catch(e){ console.log(e instanceof SyntaxError); }"#,
        ["true"]
    };

    duplicate_arrow_params_syntax_error => {
        r#"try{ eval("const d2p = (a, a) => a;"); console.log("ok"); }catch(e){ console.log(e instanceof SyntaxError); }"#,
        ["true"]
    };
}
