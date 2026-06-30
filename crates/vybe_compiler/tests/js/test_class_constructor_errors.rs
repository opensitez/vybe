//! Class constructor error paths — throw, return, super/this rules, invalid extends.

crate::js_cases! {
    constructor_throw_propagates_from_new_expression => {
        r#"try{new (class{constructor(){throw new Error("ctor");}})();}catch(e){console.log(e.message);}"#,
        ["ctor"]
    };

    constructor_throw_caught_by_surrounding_try => {
        r#"let o="";try{try{new (class{constructor(){throw "abort";}})();}catch(e){o=String(e);}}catch{o="outer";}console.log(o);"#,
        ["abort"]
    };

    constructor_return_primitive_ignored_instance_still_created => {
        r#"const C=class{constructor(){return 42;}};const i=new C();console.log(i instanceof C);console.log(typeof i);"#,
        ["true", "object"]
    };

    constructor_return_object_replaces_instance => {
        r#"const repl={tag:"repl"};const C=class{constructor(){return repl;}};const i=new C();console.log(i===repl);"#,
        ["true"]
    };

    constructor_return_null_keeps_default_instance => {
        r#"const C=class{constructor(){return null;}};const i=new C();console.log(i instanceof C);"#,
        ["true"]
    };

    derived_access_this_before_super_throws => {
        r#"try{class B{} class D extends B{constructor(){this.x=1;super();}} new D();}catch(e){console.log(e instanceof ReferenceError);}"#,
        ["true"]
    };

    derived_missing_super_call_throws_on_construct => {
        r#"try{class B{} class D extends B{constructor(){}} new D();}catch(e){console.log(e instanceof ReferenceError);}"#,
        ["true"]
    };

    derived_super_must_be_called_once => {
        r#"try{class B{} class D extends B{constructor(){super();super();}} new D();}catch(e){console.log(e instanceof ReferenceError);}"#,
        ["true"]
    };

    derived_super_with_args_passes_to_base => {
        r#"class B{constructor(v){this.base=v;}} class D extends B{constructor(v){super(v*2);}} const d=new D(3);console.log(d.base);"#,
        ["6"]
    };

    class_extends_non_object_throws => {
        r#"try{class D extends null{}} catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    class_extends_undefined_throws => {
        r#"try{class D extends undefined{}} catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    new_target_in_base_constructor => {
        r#"class C{constructor(){console.log(new.target.name);}} class D extends C{} new D();"#,
        ["D"]
    };

    new_target_undefined_on_direct_call => {
        r#"class C{constructor(){console.log(new.target===undefined);}} C();"#,
        ["true"]
    };

    constructor_called_without_new_on_class_throws => {
        r#"class C{constructor(){}} try{C();}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    arrow_not_valid_as_constructor => {
        r#"const A=class{}; try{new (()=>{})();}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    base_constructor_throw_prevents_derived_field_init => {
        r#"try{class B{constructor(){throw new Error("base");}} class D extends B{constructor(){super();this.d=1;}} new D();}catch(e){console.log(e.message);}"#,
        ["base"]
    };

    try_catch_inside_constructor_swallows_throw => {
        r#"class C{constructor(){try{throw new Error("in");}catch(e){this.msg=e.message;}}} const c=new C();console.log(c.msg);"#,
        ["in"]
    };

    constructor_rethrow_from_catch => {
        r#"try{class C{constructor(){try{throw new Error("a");}catch(e){throw e;}}} new C();}catch(e){console.log(e.message);}"#,
        ["a"]
    };

    finally_in_constructor_runs_before_throw_escapes => {
        r#"let o=[];try{class C{constructor(){try{throw 1;}finally{o.push("f");}}}} new C();}catch{o.push("c");}console.log(o.join(","));"#,
        ["f,c"]
    };

    static_block_throw_prevents_class_init => {
        r#"try{class C{static{throw new Error("static");}}} catch(e){console.log(e.message);}"#,
        ["static"]
    };

    static_block_runs_before_first_new => {
        r#"let n=0; class C{static{n++;}} new C();console.log(n);"#,
        ["1"]
    };

    field_initializer_throw_prevents_instance => {
        r#"try{class C{x=(()=>{throw new Error("init");})();}} new C();}catch(e){console.log(e.message);}"#,
        ["init"]
    };

    private_field_initializer_throw => {
        r#"try{class C{#x=(()=>{throw new Error("priv");})();}} new C();}catch(e){console.log(e.message);}"#,
        ["priv"]
    };

    extends_expression_evaluated_throw => {
        r#"function boom(){throw new Error("ext");} try{class D extends boom(){}} catch(e){console.log(e.message);}"#,
        ["ext"]
    };

    constructor_name_on_anonymous_class_expression => {
        r#"const Named=class{constructor(){console.log(Named.name);}}; new Named();"#,
        ["Named"]
    };

    derived_instanceof_base_after_successful_construct => {
        r#"class B{} class D extends B{} const d=new D();console.log(d instanceof B);console.log(d instanceof D);"#,
        ["true", "true"]
    };

    constructor_return_existing_instance_of_same_class => {
        r#"class C{constructor(){if(C.cache)return C.cache;C.cache=this;}} C.cache=null; const a=new C(); const b=new C();console.log(a===b);"#,
        ["true"]
    };

    abstract_class_cannot_instantiate_if_runtime_checks => {
        r#"class Abstract{constructor(){if(new.target===Abstract)throw new Error("abstract");}} try{new Abstract();}catch(e){console.log(e.message);}"#,
        ["abstract"]
    };

    constructor_parameter_default_throw_on_eval => {
        r#"try{class C{constructor(x=(()=>{throw new Error("def");})()){}}} new C();}catch(e){console.log(e.message);}"#,
        ["def"]
    };

    super_call_in_try_catch_in_derived => {
        r#"class B{constructor(){throw new Error("b");}} class D extends B{constructor(){try{super();}catch(e){this.recovered=e.message;}}} const d=new D();console.log(d.recovered);"#,
        ["b"]
    };

    new_on_bound_class_constructor => {
        r#"class C{constructor(v){this.v=v;}} const B=C.bind(null,7); const i=new B();console.log(i.v);"#,
        ["7"]
    };

    class_expression_inner_name_not_outer_binding => {
        r#"const Outer=class Inner{constructor(){console.log(Inner.name);}}; new Outer();"#,
        ["Inner"]
    };

    constructor_throw_typeerror_subclass => {
        r#"try{new (class{constructor(){throw new TypeError("bad");}})();}catch(e){console.log(e.name);}"#,
        ["TypeError"]
    };

    constructor_throw_aggregate_error => {
        r#"try{new (class{constructor(){throw new AggregateError([new Error("a")],"many");}})();}catch(e){console.log(e instanceof AggregateError);}"#,
        ["true"]
    };

    base_constructor_returns_object_derived_gets_it => {
        r#"const repl={x:1}; class B{constructor(){return repl;}} class D extends B{constructor(){const o=super();console.log(o===repl);}} new D();"#,
        ["true"]
    };

    duplicate_instance_field_declarations_syntax_error => {
        r#"try{eval("class C { x = 1; x = 2; }"); console.log("ok");}catch(e){console.log(e instanceof SyntaxError);}"#,
        ["true"]
    };

    constructor_with_rest_param_receives_args => {
        r#"class C{constructor(first,...rest){console.log(first);console.log(rest.join(","));}} new C(1,2,3);"#,
        ["1", "2,3"]
    };

    new_target_chain_through_super => {
        r#"class A{constructor(){this.chain=new.target.name;}} class B extends A{constructor(){super();}} class C extends B{constructor(){super();}} const c=new C();console.log(c.chain);"#,
        ["C"]
    };
}
