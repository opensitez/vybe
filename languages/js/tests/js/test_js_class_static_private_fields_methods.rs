use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Class Static Private Fields & Static Private Methods
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_class_static_private_field_access() {
    let src = r#"
class GlobalConfig {
    static #apiKey = "SECRET_123";
    static getApiKey() {
        return GlobalConfig.#apiKey;
    }
}
console.log(GlobalConfig.getApiKey());
"#;
    assert_eq!(run_js(src), vec!["SECRET_123"]);
}

#[test]
fn test_js_class_static_private_method_access() {
    let src = r#"
class Logger {
    static #format(msg) {
        return `[LOG] ${msg}`;
    }
    static log(msg) {
        return Logger.#format(msg);
    }
}
console.log(Logger.log("App Started"));
"#;
    assert_eq!(run_js(src), vec!["[LOG] App Started"]);
}

#[test]
fn test_js_class_static_private_field_outside_access_throws() {
    let src = r#"
class Vault {
    static #pass = 9999;
}
try {
    eval("Vault.#pass");
} catch (e) {
    console.log("Outside Static Private Access Error");
}
"#;
    assert_eq!(run_js(src), vec!["Outside Static Private Access Error"]);
}

#[test]
fn test_js_class_static_private_field_this_receiver_check() {
    let src = r#"
class Parent {
    static #secret = "ParentSecret";
    static getSecret() {
        return this.#secret; // 'this' must be Parent class constructor!
    }
}
class Child extends Parent {}

console.log(Parent.getSecret());
try {
    Child.getSecret(); // Called on Child constructor where #secret does NOT exist -> throws TypeError!
} catch (e) {
    console.log("Child Subclass Static Private Call Error");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["ParentSecret", "Child Subclass Static Private Call Error"]
    );
}

#[test]
fn test_js_class_static_private_getter_setter() {
    let src = r#"
class State {
    static #count = 0;
    static get #counter() { return State.#count; }
    static set #counter(v) { State.#count = v; }

    static increment() {
        State.#counter = State.#counter + 1;
    }
    static getVal() { return State.#counter; }
}
State.increment();
State.increment();
console.log(State.getVal());
"#;
    assert_eq!(run_js(src), vec!["2"]);
}

#[test]
fn test_js_class_static_private_field_initialization_order() {
    let src = r#"
let order = [];
class OrderTracker {
    static #a = (() => { order.push("Init A"); return 1; })();
    static #b = (() => { order.push("Init B"); return 2; })();
    static run() { return OrderTracker.#a + OrderTracker.#b; }
}
console.log(OrderTracker.run() + "|Order=" + order.join(","));
"#;
    assert_eq!(run_js(src), vec!["3|Order=Init A,Init B"]);
}

#[test]
fn test_js_class_static_private_field_instance_cannot_access() {
    let src = r#"
class Item {
    static #secret = 42;
    getSecretFromInstance() {
        try {
            return this.#secret;
        } catch (e) {
            return "Instance Access Error";
        }
    }
}
console.log(new Item().getSecretFromInstance());
"#;
    assert_eq!(run_js(src), vec!["Instance Access Error"]);
}

#[test]
fn test_js_class_static_private_method_subclass_shadowing() {
    let src = r#"
class Base {
    static #fn() { return "BaseStatic"; }
    static callBase() { return Base.#fn(); }
}
class Sub extends Base {
    static #fn() { return "SubStatic"; }
    static callSub() { return Sub.#fn(); }
}
console.log(`${Base.callBase()}|${Sub.callSub()}`);
"#;
    assert_eq!(run_js(src), vec!["BaseStatic|SubStatic"]);
}

#[test]
fn test_js_class_static_private_field_uninitialized_defaults_to_undefined() {
    let src = r#"
class Holder {
    static #data;
    static check() { return Holder.#data === undefined; }
}
console.log(Holder.check());
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_class_static_private_method_async() {
    let src = r#"
class AsyncLoader {
    static async #loadRaw() {
        const val = await Promise.resolve("AsyncData");
        return val;
    }
    static async fetch() {
        return await AsyncLoader.#loadRaw();
    }
}
AsyncLoader.fetch().then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["AsyncData"]);
}

#[test]
fn test_js_class_static_private_method_generator() {
    let src = r#"
class Sequence {
    static *#numbers() {
        yield 10;
        yield 20;
    }
    static getArray() {
        return [...Sequence.#numbers()].join(",");
    }
}
console.log(Sequence.getArray());
"#;
    assert_eq!(run_js(src), vec!["10,20"]);
}

#[test]
fn test_js_class_static_private_field_nullish_coalescing_assignment() {
    let src = r#"
class DB {
    static #connection = null;
    static connect() {
        DB.#connection ??= "ActiveConnection";
        return DB.#connection;
    }
}
console.log(DB.connect());
"#;
    assert_eq!(run_js(src), vec!["ActiveConnection"]);
}

#[test]
fn test_js_class_static_private_method_with_default_params() {
    let src = r#"
class Formatter {
    static #prefix(msg, tag = "[APP]") {
        return `${tag} ${msg}`;
    }
    static format(msg) { return Formatter.#prefix(msg); }
}
console.log(Formatter.format("Loaded"));
"#;
    assert_eq!(run_js(src), vec!["[APP] Loaded"]);
}

#[test]
fn test_js_class_static_private_field_cross_method_mutation() {
    let src = r#"
class Counter {
    static #val = 10;
    static add(n) { Counter.#val += n; }
    static sub(n) { Counter.#val -= n; }
    static get() { return Counter.#val; }
}
Counter.add(5);
Counter.sub(2);
console.log(Counter.get());
"#;
    assert_eq!(run_js(src), vec!["13"]);
}

#[test]
fn test_js_class_static_private_field_not_in_keys_or_get_own_property_names() {
    let src = r#"
class Sample {
    static #priv = 1;
    static pub = 2;
}
console.log(Object.keys(Sample).join(",") + "|Count=" + Object.getOwnPropertyNames(Sample).length);
"#;
    assert_eq!(run_js(src), vec!["pub|Count=4"]); // length, name, prototype, pub
}

#[test]
fn test_js_class_static_private_field_arrow_function_closure() {
    let src = r#"
class KeyStore {
    static #key = "SUPER_SECRET";
    static getGetter() {
        return () => KeyStore.#key;
    }
}
const getter = KeyStore.getGetter();
console.log(getter());
"#;
    assert_eq!(run_js(src), vec!["SUPER_SECRET"]);
}

#[test]
fn test_js_class_static_private_field_redefinition_throws_syntaxerror() {
    let src = r#"
try {
    eval("class Bad { static #x = 1; static #x = 2; }");
} catch (e) {
    console.log("Duplicate Static Private Field SyntaxError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Duplicate Static Private Field SyntaxError"]
    );
}

#[test]
fn test_js_class_static_private_field_delete_throws_syntaxerror() {
    let src = r#"
class Target {
    static #field = 10;
    static tryDelete() {
        try {
            eval("delete Target.#field;");
        } catch (e) {
            console.log("Delete Static Private Error");
        }
    }
}
Target.tryDelete();
"#;
    assert_eq!(run_js(src), vec!["Delete Static Private Error"]);
}

#[test]
fn test_js_class_static_private_field_same_name_as_instance_private_field() {
    let src = r#"
class Hybrid {
    #val = "InstancePrivate";
    static #val = "StaticPrivate";

    getInstanceVal() { return this.#val; }
    static getStaticVal() { return Hybrid.#val; }
}
const h = new Hybrid();
console.log(`${h.getInstanceVal()}|${Hybrid.getStaticVal()}`);
"#;
    assert_eq!(run_js(src), vec!["InstancePrivate|StaticPrivate"]);
}

#[test]
fn test_js_class_static_private_method_name_property() {
    let src = r#"
class Metadata {
    static #internalMethod() {}
    static getName() { return Metadata.#internalMethod.name; }
}
console.log(Metadata.getName());
"#;
    assert_eq!(run_js(src), vec!["#internalMethod"]);
}

#[test]
fn test_js_class_static_private_brand_check_subclass_returns_false() {
    let src = r#"
class Parent {
    static #brand = true;
    static isParent(target) {
        return #brand in target;
    }
}
class Child extends Parent {}
console.log(Parent.isParent(Parent) + "|" + Parent.isParent(Child));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

