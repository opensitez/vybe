use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Property Accessors (`get`, `set`) & Prototype Inheritance
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_getter_setter_object_literal_definition() {
    let src = r#"
const obj = {
    _age: 20,
    get age() { return this._age; },
    set age(v) { this._age = v; }
};
obj.age = 25;
console.log(obj.age);
"#;
    assert_eq!(run_js(src), vec!["25"]);
}

#[test]
fn test_js_getter_without_setter_ignores_assignment_in_non_strict() {
    let src = r#"
const obj = {
    get readOnly() { return "fixed"; }
};
obj.readOnly = "newVal";
console.log(obj.readOnly);
"#;
    assert_eq!(run_js(src), vec!["fixed"]);
}

#[test]
fn test_js_getter_without_setter_throws_typeerror_in_strict_mode() {
    let src = r#"
const obj = {
    get readOnly() { return "fixed"; }
};
try {
    eval("'use strict'; obj.readOnly = 'newVal';");
} catch (e) {
    console.log("ReadOnly Getter TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["ReadOnly Getter TypeError"]);
}

#[test]
fn test_js_inherited_getter_evaluates_with_receiver_this() {
    let src = r#"
const proto = {
    get value() { return this._val * 2; }
};
const obj = Object.create(proto);
obj._val = 10;
console.log(obj.value);
"#;
    assert_eq!(run_js(src), vec!["20"]);
}

#[test]
fn test_js_inherited_setter_invokes_with_receiver_this() {
    let src = r#"
const proto = {
    set value(v) { this._val = v + 1; }
};
const obj = Object.create(proto);
obj.value = 5;
console.log(obj._val + "|hasOwn=" + Object.hasOwn(obj, "_val"));
"#;
    assert_eq!(run_js(src), vec!["6|hasOwn=true"]); // Setter creates property on receiver 'obj', not 'proto'!
}

#[test]
fn test_js_getter_setter_class_prototype_inheritance() {
    let src = r#"
class Circle {
    constructor(radius) { this.radius = radius; }
    get area() { return Math.PI * this.radius ** 2; }
}
const c = new Circle(2);
console.log(c.area.toFixed(2));
"#;
    assert_eq!(run_js(src), vec!["12.57"]);
}

#[test]
fn test_js_computed_getter_setter_names() {
    let src = r#"
const propName = "dynamicProp";
const obj = {
    _val: "init",
    get [propName]() { return this._val; },
    set [propName](v) { this._val = v; }
};
obj.dynamicProp = "updated";
console.log(obj.dynamicProp);
"#;
    assert_eq!(run_js(src), vec!["updated"]);
}

#[test]
fn test_js_symbol_computed_getter_setter_names() {
    let src = r#"
const sym = Symbol("privateVal");
const obj = {
    _val: 100,
    get [sym]() { return this._val; },
    set [sym](v) { this._val = v; }
};
obj[sym] = 200;
console.log(obj[sym]);
"#;
    assert_eq!(run_js(src), vec!["200"]);
}

#[test]
fn test_js_getter_throwing_exception() {
    let src = r#"
const obj = {
    get fail() { throw new Error("GetterFailed"); }
};
try {
    const val = obj.fail;
} catch (e) {
    console.log(e.message);
}
"#;
    assert_eq!(run_js(src), vec!["GetterFailed"]);
}

#[test]
fn test_js_setter_throwing_exception() {
    let src = r#"
const obj = {
    set fail(v) { throw new Error("SetterFailed"); }
};
try {
    obj.fail = 10;
} catch (e) {
    console.log(e.message);
}
"#;
    assert_eq!(run_js(src), vec!["SetterFailed"]);
}

#[test]
fn test_js_getter_setter_descriptors_defined_via_define_property() {
    let src = r#"
const obj = { _val: 0 };
Object.defineProperty(obj, "val", {
    get() { return this._val; },
    set(v) { this._val = v * 10; },
    enumerable: true,
    configurable: true
});
obj.val = 5;
console.log(obj.val);
"#;
    assert_eq!(run_js(src), vec!["50"]);
}

#[test]
fn test_js_getter_setter_property_descriptor_structure() {
    let src = r#"
const obj = {
    get x() { return 1; }
};
const desc = Object.getOwnPropertyDescriptor(obj, "x");
console.log(`${typeof desc.get}:${desc.set}:${desc.value}:${desc.writable}`);
"#;
    assert_eq!(run_js(src), vec!["function:undefined:undefined:undefined"]);
}

#[test]
fn test_js_overriding_inherited_getter_with_own_data_property() {
    let src = r#"
const proto = {
    get name() { return "ProtoName"; }
};
const obj = Object.create(proto);
obj.name = "OwnName"; // Shadowing getter with own property fails in non-strict mode if non-writable in proto!
console.log(obj.name);
"#;
    assert_eq!(run_js(src), vec!["ProtoName"]);
}

#[test]
fn test_js_shadowing_inherited_getter_via_define_property() {
    let src = r#"
const proto = {
    get name() { return "ProtoName"; }
};
const obj = Object.create(proto);
Object.defineProperty(obj, "name", { value: "OwnName", writable: true });
console.log(obj.name);
"#;
    assert_eq!(run_js(src), vec!["OwnName"]);
}

#[test]
fn test_js_static_class_getter_setter() {
    let src = r#"
class Config {
    static _env = "dev";
    static get env() { return this._env; }
    static set env(v) { this._env = v; }
}
Config.env = "prod";
console.log(Config.env);
"#;
    assert_eq!(run_js(src), vec!["prod"]);
}

#[test]
fn test_js_getter_setter_in_object_assign_evaluated() {
    let src = r#"
const source = {
    get val() { return "EvaluatedValue"; }
};
const target = Object.assign({}, source);
const desc = Object.getOwnPropertyDescriptor(target, "val");
console.log(target.val + "|isDataProperty=" + (desc.get === undefined));
"#;
    assert_eq!(run_js(src), vec!["EvaluatedValue|isDataProperty=true"]); // Object.assign copies evaluated values as data properties!
}

#[test]
fn test_js_getter_deleting_self_property() {
    let src = r#"
const obj = {
    get temp() {
        delete this.temp;
        return (this.temp = "CachedVal");
    }
};
console.log(obj.temp + "|" + obj.temp);
"#;
    assert_eq!(run_js(src), vec!["CachedVal|CachedVal"]);
}

#[test]
fn test_js_setter_accepts_exactly_one_argument() {
    let src = r#"
const obj = {
    set val(...args) {
        console.log(args.length + "|" + args[0]);
    }
};
obj.val = 42;
"#;
    assert_eq!(run_js(src), vec!["1|42"]);
}

#[test]
fn test_js_getter_takes_no_arguments() {
    let src = r#"
const obj = {
    get fn() { return arguments ? arguments.length : 0; }
};
console.log(obj.fn);
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_super_getter_call_in_derived_class() {
    let src = r#"
class Parent {
    get label() { return "ParentLabel"; }
}
class Child extends Parent {
    get label() { return super.label + " -> ChildLabel"; }
}
const c = new Child();
console.log(c.label);
"#;
    assert_eq!(run_js(src), vec!["ParentLabel -> ChildLabel"]);
}
