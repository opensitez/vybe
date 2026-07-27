use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `super.method()` & `super.property` Prototype Access Mechanics
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_class_super_method_call_overriding() {
    let src = r#"
class Animal {
    makeSound() { return "Generic Sound"; }
}
class Dog extends Animal {
    makeSound() {
        return super.makeSound() + " -> Woof";
    }
}
console.log(new Dog().makeSound());
"#;
    assert_eq!(run_js(src), vec!["Generic Sound -> Woof"]);
}

#[test]
fn test_js_class_super_method_this_receiver_preservation() {
    let src = r#"
class Parent {
    getName() { return this.name; }
}
class Child extends Parent {
    constructor(name) {
        super();
        this.name = name;
    }
    getName() {
        return super.getName().toUpperCase(); // super.getName() executes with 'this' pointing to Child instance!
    }
}
console.log(new Child("alice").getName());
"#;
    assert_eq!(run_js(src), vec!["ALICE"]);
}

#[test]
fn test_js_class_super_property_getter_access() {
    let src = r#"
class Base {
    get value() { return 100; }
}
class Derived extends Base {
    get value() {
        return super.value * 2;
    }
}
console.log(new Derived().value);
"#;
    assert_eq!(run_js(src), vec!["200"]);
}

#[test]
fn test_js_class_super_property_setter_assignment() {
    let src = r#"
class Base {
    set data(v) { this._data = v + 10; }
    get data() { return this._data; }
}
class Derived extends Base {
    setData(v) {
        super.data = v; // super.data = v sets property on 'this' receiver using Base's setter!
    }
}
const d = new Derived();
d.setData(50);
console.log(d.data);
"#;
    assert_eq!(run_js(src), vec!["60"]);
}

#[test]
fn test_js_class_super_static_method_call() {
    let src = r#"
class Logger {
    static log(msg) { return `[LOG] ${msg}`; }
}
class CustomLogger extends Logger {
    static log(msg) {
        return super.log(msg).toUpperCase();
    }
}
console.log(CustomLogger.log("Hello"));
"#;
    assert_eq!(run_js(src), vec!["[LOG] HELLO"]);
}

#[test]
fn test_js_class_super_static_property_chain() {
    let src = r#"
class Base {
    static get nameTag() { return "Base"; }
}
class Sub extends Base {
    static get nameTag() { return super.nameTag + "->Sub"; }
}
console.log(Sub.nameTag);
"#;
    assert_eq!(run_js(src), vec!["Base->Sub"]);
}

#[test]
fn test_js_class_super_static_setter_assignment() {
    assert_eq!(
        run_js(
            r#"
class Base {
    static set marker(v) {
        this._marker = `base:${v}`;
    }
    static get marker() {
        return this._marker;
    }
}
class Derived extends Base {
    static applyMarker(v) {
        super.marker = v;
    }
}

Derived.applyMarker("X");
console.log(Derived.marker);
console.log(Base.marker);
console.log(Object.hasOwn(Derived, "_marker"));
console.log(Object.hasOwn(Base, "_marker"));
"#,
        ),
        vec!["base:X", "undefined", "true", "false"]
    );
}

#[test]
fn test_js_class_super_method_call_in_object_literal_concise_method() {
    let src = r#"
const parent = {
    greet() { return "ParentGreet"; }
};
const child = {
    __proto__: parent,
    greet() {
        return super.greet() + "Child";
    }
};
console.log(child.greet());
"#;
    assert_eq!(run_js(src), vec!["ParentGreetChild"]);
}

#[test]
fn test_js_class_super_access_in_arrow_function_inside_method() {
    let src = r#"
class Base {
    fetch() { return "BaseData"; }
}
class Sub extends Base {
    getFetcher() {
        return () => super.fetch(); // Arrow function inherits HomeObject for super!
    }
}
const s = new Sub();
const fetcher = s.getFetcher();
console.log(fetcher());
"#;
    assert_eq!(run_js(src), vec!["BaseData"]);
}

#[test]
fn test_js_class_super_property_read_own_instance_field_returns_undefined() {
    let src = r#"
class Base {
    baseField = "InstanceField";
}
class Sub extends Base {
    checkSuperField() {
        return super.baseField === undefined; // super looks up on Prototype chain, NOT instance fields!
    }
}
console.log(new Sub().checkSuperField());
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_class_super_method_chained_three_levels() {
    let src = r#"
class A { test() { return "A"; } }
class B extends A { test() { return super.test() + "B"; } }
class C extends B { test() { return super.test() + "C"; } }

console.log(new C().test());
"#;
    assert_eq!(run_js(src), vec!["ABC"]);
}

#[test]
fn test_js_class_super_method_call_with_arguments_object() {
    let src = r#"
class Base {
    add(a, b) { return a + b; }
}
class Sub extends Base {
    add() {
        return super.add(arguments[0], arguments[1]) * 10;
    }
}
console.log(new Sub().add(2, 3));
"#;
    assert_eq!(run_js(src), vec!["50"]);
}

#[test]
fn test_js_class_super_property_delete_throws_referenceerror() {
    let src = r#"
class Base {
    foo() {}
}
class Sub extends Base {
    deleteFoo() {
        try {
            eval("delete super.foo;");
        } catch (e) {
            console.log("Delete Super ReferenceError");
        }
    }
}
new Sub().deleteFoo();
"#;
    assert_eq!(run_js(src), vec!["Delete Super ReferenceError"]);
}

#[test]
fn test_js_class_super_call_outside_class_or_concise_method_throws() {
    let src = r#"
try {
    eval("function standalone() { super.foo(); } standalone();");
} catch (e) {
    console.log("Super Outside Method Error");
}
"#;
    assert_eq!(run_js(src), vec!["Super Outside Method Error"]);
}

#[test]
fn test_js_class_super_method_generator_yield_star() {
    let src = r#"
class Base {
    *generate() {
        yield 1; yield 2;
    }
}
class Sub extends Base {
    *generate() {
        yield* super.generate();
        yield 3;
    }
}
console.log([...new Sub().generate()].join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2,3"]);
}

#[test]
fn test_js_class_super_method_async_await() {
    let src = r#"
class Base {
    async load() { return await Promise.resolve("BaseLoad"); }
}
class Sub extends Base {
    async load() {
        const val = await super.load();
        return val + "_Extended";
    }
}
new Sub().load().then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["BaseLoad_Extended"]);
}

#[test]
fn test_js_class_super_method_bound_function() {
    let src = r#"
class Base {
    compute(x) { return x + 10; }
}
class Sub extends Base {
    getBoundCompute() {
        return super.compute.bind(this);
    }
}
const s = new Sub();
const bound = s.getBoundCompute();
console.log(bound(5));
"#;
    assert_eq!(run_js(src), vec!["15"]);
}

#[test]
fn test_js_class_super_property_symbol_key_access() {
    let src = r#"
const sym = Symbol("action");
class Base {
    [sym]() { return "SymbolBase"; }
}
class Sub extends Base {
    [sym]() {
        return super[sym]() + "_Sub";
    }
}
console.log(new Sub()[sym]());
"#;
    assert_eq!(run_js(src), vec!["SymbolBase_Sub"]);
}

#[test]
fn test_js_class_super_method_rebound_prototype_chain() {
    let src = r#"
class Base1 { action() { return "B1"; } }
class Base2 { action() { return "B2"; } }
class Sub extends Base1 {
    action() { return super.action(); }
}
const s = new Sub();
Object.setPrototypeOf(Sub.prototype, Base2.prototype);
console.log(s.action()); // HomeObject[[Prototype]] is Base2 -> returns "B2"!
"#;
    assert_eq!(run_js(src), vec!["B2"]);
}

#[test]
fn test_js_class_super_property_increment_decrement_operators() {
    let src = r#"
class Base {
    get count() { return this._c || 0; }
    set count(v) { this._c = v; }
}
class Sub extends Base {
    increment() {
        super.count++; // Performs super.count = super.count + 1 on 'this'
    }
}
const s = new Sub();
s.increment();
s.increment();
console.log(s.count);
"#;
    assert_eq!(run_js(src), vec!["2"]);
}

#[test]
fn test_js_class_super_property_in_constructor() {
    let src = r#"
class Base {
    initMessage() { return "Initialized"; }
}
class Sub extends Base {
    constructor() {
        super();
        this.msg = super.initMessage();
    }
}
console.log(new Sub().msg);
"#;
    assert_eq!(run_js(src), vec!["Initialized"]);
}

#[test]
fn test_js_class_super_method_apply_call_receiver() {
    let src = r#"
class Base {
    greet() { return `Hello ${this.name}`; }
}
class Sub extends Base {
    greet() {
        return super.greet.call({ name: "CustomReceiver" });
    }
}
console.log(new Sub().greet());
"#;
    assert_eq!(run_js(src), vec!["Hello CustomReceiver"]);
}
