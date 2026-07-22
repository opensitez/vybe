use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Private `#field in object` Brand Check Operator (ES2022)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_private_in_operator_instance_brand_check() {
    let src = r#"
class Person {
    #name;
    constructor(n) { this.#name = n; }

    static isPerson(obj) {
        return #name in obj;
    }
}
const p = new Person("Alice");
console.log(Person.isPerson(p) + "|" + Person.isPerson({}));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_private_in_operator_primitive_returns_false() {
    let src = r#"
class Box {
    #content;
    static isBox(val) {
        return #content in val;
    }
}
console.log(Box.isBox(42) + "|" + Box.isBox("str") + "|" + Box.isBox(null));
"#;
    assert_eq!(run_js(src), vec!["false|false|false"]);
}

#[test]
fn test_js_private_in_operator_private_method_brand_check() {
    let src = r#"
class Validator {
    #validate() {}
    static check(obj) {
        return #validate in obj;
    }
}
const v = new Validator();
console.log(Validator.check(v) + "|" + Validator.check({}));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_private_in_operator_private_getter_brand_check() {
    let src = r#"
class User {
    get #secret() { return 42; }
    static hasSecret(obj) {
        return #secret in obj;
    }
}
console.log(User.hasSecret(new User()) + "|" + User.hasSecret(Object.create(new User())));
"#;
    assert_eq!(run_js(src), vec!["true|false"]); // Object.create(instance) does NOT have private field brand!
}

#[test]
fn test_js_private_in_operator_static_private_brand_check() {
    let src = r#"
class System {
    static #config = {};
    static isSystemClass(target) {
        return #config in target;
    }
}
console.log(System.isSystemClass(System) + "|" + System.isSystemClass(new System()));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_private_in_operator_inherited_prototype_returns_false() {
    let src = r#"
class Parent {
    #parentField = 100;
    static hasParentField(obj) {
        return #parentField in obj;
    }
}
class Child extends Parent {}

const p = new Parent();
const c = new Child();
console.log(Parent.hasParentField(p) + "|" + Parent.hasParentField(c));
"#;
    assert_eq!(run_js(src), vec!["true|true"]); // Child inherits instance private fields via super()!
}

#[test]
fn test_js_private_in_operator_subclass_without_super_returns_false() {
    let src = r#"
class Base {
    #baseField = 10;
    static check(obj) { return #baseField in obj; }
}
const plainObj = {};
console.log(Base.check(plainObj));
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_private_in_operator_inside_instance_method() {
    let src = r#"
class Node {
    #val;
    constructor(v) { this.#val = v; }
    isSameClass(other) {
        return #val in other;
    }
}
const n1 = new Node(1);
const n2 = new Node(2);
console.log(n1.isSameClass(n2) + "|" + n1.isSameClass({}));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_private_in_operator_undeclared_private_name_throws_syntaxerror() {
    let src = r#"
try {
    eval("class Test { check(o) { return #undeclared in o; } }");
} catch (e) {
    console.log("Undeclared Private In SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Undeclared Private In SyntaxError"]);
}

#[test]
fn test_js_private_in_operator_subclass_static_private_brand_check() {
    let src = r#"
class Parent {
    static #parentStatic = 1;
    static isParentClass(target) {
        return #parentStatic in target;
    }
}
class Child extends Parent {}

console.log(Parent.isParentClass(Parent) + "|" + Parent.isParentClass(Child));
"#;
    assert_eq!(run_js(src), vec!["true|false"]); // Child class constructor does NOT possess Parent's static private brand!
}

#[test]
fn test_js_private_in_operator_shadowed_private_name_in_subclass() {
    let src = r#"
class Base {
    #tag = "Base";
    static isBase(o) { return #tag in o; }
}
class Sub extends Base {
    #tag = "Sub";
    static isSub(o) { return #tag in o; }
}
const b = new Base();
const s = new Sub();
console.log(`${Base.isBase(b)}|${Base.isBase(s)}|${Sub.isSub(b)}|${Sub.isSub(s)}`);
"#;
    assert_eq!(run_js(src), vec!["true|true|false|true"]);
}

#[test]
fn test_js_private_in_operator_proxy_target_inspection() {
    let src = r#"
class Token {
    #tokenVal = "ABC";
    static isToken(obj) { return #tokenVal in obj; }
}
const t = new Token();
const proxy = new Proxy(t, {});
console.log(Token.isToken(t) + "|" + Token.isToken(proxy)); // Brand check on Proxy succeeds if target is instance
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_private_in_operator_with_null_and_undefined() {
    let src = r#"
class Guard {
    #field;
    static safeCheck(o) {
        return #field in o;
    }
}
console.log(Guard.safeCheck(null) + "|" + Guard.safeCheck(undefined));
"#;
    assert_eq!(run_js(src), vec!["false|false"]);
}

#[test]
fn test_js_private_in_operator_ternary_condition() {
    let src = r#"
class Account {
    #balance = 100;
    static read(obj) {
        return #balance in obj ? obj.#balance : -1;
    }
}
const a = new Account();
console.log(Account.read(a) + "|" + Account.read({}));
"#;
    assert_eq!(run_js(src), vec!["100|-1"]);
}

#[test]
fn test_js_private_in_operator_uninitialized_field_brand_check() {
    let src = r#"
class EmptyHolder {
    #empty;
    static hasBrand(o) { return #empty in o; }
}
const h = new EmptyHolder();
console.log(EmptyHolder.hasBrand(h)); // Brand check is true even if field is undefined!
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_private_in_operator_multiple_fields_brand_check() {
    let src = r#"
class Entity {
    #id; #name;
    static isEntity(o) {
        return #id in o && #name in o;
    }
}
console.log(Entity.isEntity(new Entity()));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_private_in_operator_negation() {
    let src = r#"
class Secret {
    #code;
    static isNotSecret(o) {
        return !(#code in o);
    }
}
console.log(Secret.isNotSecret({}) + "|" + Secret.isNotSecret(new Secret()));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_private_in_operator_in_loop_filtering() {
    let src = r#"
class Target {
    #id = 1;
    static countValid(arr) {
        return arr.filter(item => #id in item).length;
    }
}
const items = [new Target(), {}, new Target(), "str"];
console.log(Target.countValid(items));
"#;
    assert_eq!(run_js(src), vec!["2"]);
}

#[test]
fn test_js_private_in_operator_arrow_function_lexical_brand_check() {
    let src = r#"
class Container {
    #val = 50;
    getChecker() {
        return o => #val in o;
    }
}
const c = new Container();
const check = c.getChecker();
console.log(check(c) + "|" + check({}));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_private_in_operator_with_object_create_null() {
    let src = r#"
class Box {
    #content;
    static check(o) { return #content in o; }
}
const nullProto = Object.create(null);
console.log(Box.check(nullProto));
"#;
    assert_eq!(run_js(src), vec!["false"]);
}
