//! Private field/method brand errors — access outside declaring class, static vs instance.

crate::js_cases! {
    private_field_read_outside_class_throws => {
        r#"class C{#x=1;get(){return this.#x;}} const c=new C(); try{console.log(c.#x);}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    private_field_write_outside_class_throws => {
        r#"class C{#x=1;} const c=new C(); try{c.#x=2;}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    private_method_call_outside_class_throws => {
        r#"class C{#m(){return 1;}} const c=new C(); try{c.#m();}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    private_getter_outside_class_throws => {
        r#"class C{get #v(){return 1;}} const c=new C(); try{console.log(c.#v);}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    private_setter_outside_class_throws => {
        r#"class C{set #v(x){} } const c=new C(); try{c.#v=1;}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    private_static_field_outside_class_throws => {
        r#"class C{static #s=1;} try{console.log(C.#s);}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    private_static_method_outside_class_throws => {
        r#"class C{static #m(){return 1;}} try{C.#m();}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    private_in_check_true_on_branded_instance => {
        r#"class C{#token=1; static isC(v){return #token in v;}} const c=new C();console.log(C.isC(c));"#,
        ["true"]
    };

    private_in_check_false_on_plain_object => {
        r#"class C{#token=1; static isC(v){return #token in v;}} console.log(C.isC({}));"#,
        ["false"]
    };

    private_in_check_false_on_different_class_instance => {
        r#"class A{#a=1; static isA(v){return #a in v;}} class B{#b=1;} console.log(A.isA(new B()));"#,
        ["false"]
    };

    subclass_cannot_access_base_private_field => {
        r#"class B{#secret=1;} class D extends B{read(){try{return this.#secret;}catch(e){return "err";}}} console.log(new D().read());"#,
        ["err"]
    };

    subclass_cannot_call_base_private_method => {
        r#"class B{#m(){return 1;}} class D extends B{call(){try{return this.#m();}catch(e){return "err";}}} console.log(new D().call());"#,
        ["err"]
    };

    private_field_access_inside_declaring_class_works => {
        r#"class C{#n=5; double(){return this.#n*2;}} console.log(new C().double());"#,
        ["10"]
    };

    private_static_access_inside_class_works => {
        r#"class C{static #v=3; static get(){return C.#v;}} console.log(C.get());"#,
        ["3"]
    };

    private_name_collision_distinct_per_class => {
        r#"class A{#x(){return "a";}} class B{#x(){return "b";}} console.log(new A().#x());console.log(new B().#x());"#,
        ["a", "b"]
    };

    // Node-verified: the original (`const {#x} = c` outside the class)
    // is an EARLY SyntaxError — unparseable, not a catchable TypeError.
    // The runtime-testable concept: reading a declared private FIELD on
    // a non-instance fails the brand check with TypeError (§8.3.6).
    destructuring_private_field_outside_class_throws => {
        r#"class C{ #x=1; static read(o){ return o.#x; } } try{ C.read({}); }catch(e){ console.log(e instanceof TypeError); }"#,
        ["true"]
    };

    private_field_in_computed_key_outside_throws => {
        r#"class C{#k=1;} const c=new C(); try{const o={[c.#k]:1};}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    private_field_via_reflect_get_still_branded => {
        r#"class C{#x=9; getX(){return this.#x;}} const c=new C();console.log(c.getX());"#,
        ["9"]
    };

    private_method_used_from_public_method => {
        r#"class C{#inc(v){return v+1;} bump(v){return this.#inc(v);}} console.log(new C().bump(4));"#,
        ["5"]
    };

    private_accessor_from_public_getter => {
        r#"class C{#v=2; get value(){return this.#v;} } console.log(new C().value);"#,
        ["2"]
    };

    private_field_initialized_from_constructor_param => {
        r#"class C{#id; constructor(id){this.#id=id;} get(){return this.#id;}} console.log(new C("abc").get());"#,
        ["abc"]
    };

    private_field_arrow_in_class_body_captures_class => {
        r#"class C{#n=1; fn=()=>this.#n;} console.log(new C().fn());"#,
        ["1"]
    };

    private_static_block_can_set_private_static => {
        r#"class C{static #v; static{ C.#v = 7; } static read(){return C.#v;}} console.log(C.read());"#,
        ["7"]
    };

    instance_cannot_read_private_static_field => {
        r#"class C{static #s=1;} const c=new C(); try{console.log(c.#s);}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    private_method_on_wrong_receiver_throws => {
        r#"class C{#m(){return 1;} run(o){return o.#m();}} const a=new C(); const b=new C(); try{a.run({});}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    private_getter_throw_inside_accessor => {
        r#"class C{get #bad(){throw new Error("priv");} read(){try{return this.#bad;}catch(e){return e.message;}}} console.log(new C().read());"#,
        ["priv"]
    };

    private_setter_throw_inside_accessor => {
        r#"class C{set #bad(v){throw new Error("set");} write(){try{this.#bad=1;}catch(e){return e.message;}}} console.log(new C().write());"#,
        ["set"]
    };

    two_instances_share_brand_not_each_others_private => {
        r#"class C{#v; constructor(v){this.#v=v;} sameBrand(other){return #v in other;}} const a=new C(1); const b=new C(2);console.log(a.sameBrand(b));"#,
        ["true"]
    };

    private_field_name_not_in_object_keys => {
        r#"class C{#hidden=1; x=2;} console.log(Object.keys(new C()).join(","));"#,
        ["x"]
    };

    private_field_not_enumerable_in_reflect_ownkeys => {
        r#"class C{#h=1; a=2;} const k=Reflect.ownKeys(new C());console.log(k.includes("a"));console.log(k.some(x=>typeof x==="symbol"));"#,
        ["true", "false"]
    };

    extends_public_method_can_use_own_private_not_base => {
        r#"class B{#b=1; getB(){return this.#b;}} class D extends B{getD(){return this.#d;} #d=2;} const d=new D();console.log(d.getD());console.log(d.getB());"#,
        ["2", "1"]
    };

    private_in_on_null_throws => {
        r#"class C{#x=1; static check(v){return #x in v;}} try{C.check(null);}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    private_in_on_primitive_throws => {
        r#"class C{#x=1; static check(v){return #x in v;}} try{C.check(1);}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    class_expression_private_field_brand => {
        r#"const K=class{#v=3; read(){return this.#v;}}; console.log(new K().read());"#,
        ["3"]
    };

    nested_class_private_not_visible_to_outer => {
        r#"class Outer{static make(){class Inner{#x=1; get(){return this.#x;}} return new Inner();}} console.log(Outer.make().get());"#,
        ["1"]
    };

    private_method_called_with_super_in_subclass_public => {
        r#"class B{#base(){return "b";} pub(){return this.#base();}} class D extends B{wrap(){return super.pub();}} console.log(new D().wrap());"#,
        ["b"]
    };

    private_in_on_undefined_throws => {
        r#"class C{#x=1; static check(v){return #x in v;}} try{C.check(undefined);}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };
}

