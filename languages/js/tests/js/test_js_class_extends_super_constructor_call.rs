use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Class `extends` & `super(...)` Constructor Semantics
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_class_extends_super_constructor_forwarding() {
    let src = r#"
class Parent {
    constructor(name) {
        this.name = name;
    }
}
class Child extends Parent {
    constructor(name, age) {
        super(name);
        this.age = age;
    }
}
const c = new Child("Alice", 10);
console.log(`${c.name}:${c.age}`);
"#;
    assert_eq!(run_js(src), vec!["Alice:10"]);
}

#[test]
fn test_js_class_derived_constructor_must_call_super_before_this() {
    let src = r#"
class Base {}
class Sub extends Base {
    constructor() {
        try {
            eval("this.x = 10; super();");
        } catch (e) {
            console.log("This Access Before Super ReferenceError");
        }
    }
}
new Sub();
"#;
    assert_eq!(run_js(src), vec!["This Access Before Super ReferenceError"]);
}

#[test]
fn test_js_class_derived_implicit_constructor_calls_super() {
    let src = r#"
class Parent {
    constructor(val) {
        this.val = val;
    }
}
class Child extends Parent {} // Implicit constructor(...args) { super(...args); }
const c = new Child("ImplicitVal");
console.log(c.val);
"#;
    assert_eq!(run_js(src), vec!["ImplicitVal"]);
}

#[test]
fn test_js_class_extends_null_prototype() {
    let src = r#"
class NullBase extends null {
    constructor() {
        return Object.create(null);
    }
}
const nb = new NullBase();
console.log(Object.getPrototypeOf(nb) === null);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_class_super_returns_custom_object_from_parent_ctor() {
    let src = r#"
class Parent {
    constructor() {
        return { customObject: true };
    }
}
class Child extends Parent {
    constructor() {
        super();
    }
}
const c = new Child();
console.log(c.customObject);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_class_extends_builtin_array_subclassing() {
    let src = r#"
class CustomArray extends Array {
    first() { return this[0]; }
    last() { return this[this.length - 1]; }
}
const arr = new CustomArray(10, 20, 30);
console.log(arr.first() + "|" + arr.last() + "|" + (arr instanceof Array));
"#;
    assert_eq!(run_js(src), vec!["10|30|true"]);
}

#[test]
fn test_js_class_extends_builtin_error_subclassing() {
    let src = r#"
class HttpError extends Error {
    constructor(status, message) {
        super(message);
        this.name = "HttpError";
        this.status = status;
    }
}
const err = new HttpError(404, "Not Found");
console.log(err.name + ":" + err.status + "|" + err.message + "|" + (err instanceof Error));
"#;
    assert_eq!(run_js(src), vec!["HttpError:404|Not Found|true"]);
}

#[test]
fn test_js_class_extends_expression_dynamic_inheritance() {
    let src = r#"
function createBaseClass(prefix) {
    return class {
        getPrefix() { return prefix; }
    };
}
class DynamicChild extends createBaseClass("[DYN]") {}
console.log(new DynamicChild().getPrefix());
"#;
    assert_eq!(run_js(src), vec!["[DYN]"]);
}

#[test]
fn test_js_class_extends_non_constructor_throws_typeerror() {
    let src = r#"
try {
    eval("class Bad extends 12345 {}");
} catch (e) {
    console.log("Extends Non-Constructor TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Extends Non-Constructor TypeError"]);
}

#[test]
fn test_js_class_instanceof_multilevel_inheritance() {
    let src = r#"
class GrandParent {}
class Parent extends GrandParent {}
class Child extends Parent {}

const c = new Child();
console.log((c instanceof Child) + "|" + (c instanceof Parent) + "|" + (c instanceof GrandParent) + "|" + (c instanceof Object));
"#;
    assert_eq!(run_js(src), vec!["true|true|true|true"]);
}

#[test]
fn test_js_class_static_inheritance_chain() {
    let src = r#"
class Parent {
    static parentStatic() { return "ParentStaticVal"; }
}
class Child extends Parent {}

console.log(Child.parentStatic());
"#;
    assert_eq!(run_js(src), vec!["ParentStaticVal"]);
}

#[test]
fn test_js_class_super_called_twice_throws_referenceerror() {
    let src = r#"
class Base {}
class Sub extends Base {
    constructor() {
        super();
        try {
            super(); // Double super call throws ReferenceError!
        } catch (e) {
            console.log("Double Super ReferenceError");
        }
    }
}
new Sub();
"#;
    assert_eq!(run_js(src), vec!["Double Super ReferenceError"]);
}

#[test]
fn test_js_class_derived_ctor_returning_primitive_returns_this() {
    let src = r#"
class Base {}
class Sub extends Base {
    constructor() {
        super();
        return 42; // Primitive return in derived constructor is ignored!
    }
}
const s = new Sub();
console.log(s instanceof Sub);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_class_derived_ctor_returning_custom_object_overrides_this() {
    let src = r#"
class Base {}
class Sub extends Base {
    constructor() {
        super();
        return { custom: "OverrideThis" };
    }
}
const s = new Sub();
console.log(s.custom + "|" + (s instanceof Sub));
"#;
    assert_eq!(run_js(src), vec!["OverrideThis|false"]);
}

#[test]
fn test_js_class_extends_new_target_is_derived() {
    let src = r#"
class Base {
    constructor() {
        this.createdBy = new.target.name;
    }
}
class Child extends Base {}
console.log(new Child().createdBy);
"#;
    assert_eq!(run_js(src), vec!["Child"]);
}

#[test]
fn test_js_class_prototype_property_descriptor_non_enumerable() {
    let src = r#"
class Foo {}
const desc = Object.getOwnPropertyDescriptor(Foo, "prototype");
console.log(desc.writable + "|" + desc.enumerable + "|" + desc.configurable);
"#;
    assert_eq!(run_js(src), vec!["false|false|false"]);
}

#[test]
fn test_js_class_extends_builtin_map_subclassing() {
    let src = r#"
class DefaultMap extends Map {
    get(key) {
        if (!this.has(key)) this.set(key, 0);
        return super.get(key);
    }
}
const m = new DefaultMap();
console.log(m.get("counter"));
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_class_constructor_name_property() {
    let src = r#"
class NamedClass {}
console.log(NamedClass.name + "|" + new NamedClass().constructor.name);
"#;
    assert_eq!(run_js(src), vec!["NamedClass|NamedClass"]);
}

#[test]
fn test_js_class_super_spread_arguments_forwarding() {
    let src = r#"
class Base {
    constructor(...args) {
        this.sum = args.reduce((a, b) => a + b, 0);
    }
}
class Sub extends Base {
    constructor(multiplier, ...nums) {
        super(...nums);
        this.total = this.sum * multiplier;
    }
}
console.log(new Sub(10, 1, 2, 3).total);
"#;
    assert_eq!(run_js(src), vec!["60"]);
}

#[test]
fn test_js_class_extends_object_behavior() {
    let src = r#"
class ExplicitObjectExtends extends Object {
    constructor(v) {
        super(v);
    }
}
const o = new ExplicitObjectExtends("test");
console.log(o.toString());
"#;
    assert_eq!(run_js(src), vec!["test"]);
}

#[test]
fn test_js_class_constructor_cannot_be_called_without_new() {
    let src = r#"
class Point {}
try {
    Point();
} catch (e) {
    console.log("Class Call Without New TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Class Call Without New TypeError"]);
}
