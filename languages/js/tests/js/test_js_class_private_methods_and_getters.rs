use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Class Private Methods & Private Accessors (#method, #get/#set)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_class_private_method_invocation() {
    let src = r#"
class Calculator {
    #double(x) { return x * 2; }
    compute(n) { return this.#double(n); }
}
console.log(new Calculator().compute(21));
"#;
    assert_eq!(run_js(src), vec!["42"]);
}

#[test]
fn test_js_class_private_getter_setter_access() {
    let src = r#"
class Temperature {
    #celsius = 0;
    get #fahrenheit() { return this.#celsius * 1.8 + 32; }
    set #fahrenheit(f) { this.#celsius = (f - 32) / 1.8; }

    setFahrenheit(f) { this.#fahrenheit = f; }
    getFahrenheit() { return this.#fahrenheit; }
}
const t = new Temperature();
t.setFahrenheit(100);
console.log(t.getFahrenheit().toFixed(1));
"#;
    assert_eq!(run_js(src), vec!["100.0"]);
}

#[test]
fn test_js_class_private_method_outside_invocation_throws_typeerror() {
    let src = r#"
class Service {
    #internalAction() { return "Internal"; }
}
const s = new Service();
try {
    eval("s.#internalAction()");
} catch (e) {
    console.log("Outside Private Method Call Error");
}
"#;
    assert_eq!(run_js(src), vec!["Outside Private Method Call Error"]);
}

#[test]
fn test_js_class_private_method_wrong_receiver_throws_typeerror() {
    let src = r#"
class Component {
    #render() { return "DOM"; }
    callRender(other) {
        return other.#render();
    }
}
const c = new Component();
try {
    c.callRender({});
} catch (e) {
    console.log("Private Method Wrong Receiver TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Private Method Wrong Receiver TypeError"]);
}

#[test]
fn test_js_class_private_generator_method() {
    let src = r#"
class Sequence {
    async *#generate() {
        yield 1;
        yield 2;
    }
    async getItems() {
        const res = [];
        for await (const x of this.#generate()) res.push(x);
        return res.join(",");
    }
}
new Sequence().getItems().then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["1,2"]);
}

#[test]
fn test_js_class_private_async_method() {
    let src = r#"
class RemoteFetcher {
    async #fetchData() {
        const val = await Promise.resolve("RemoteVal");
        return val.toUpperCase();
    }
    async load() {
        return await this.#fetchData();
    }
}
new RemoteFetcher().load().then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["REMOTEVAL"]);
}

#[test]
fn test_js_class_private_method_cross_instance_call() {
    let src = r#"
class EncryptedString {
    #data;
    constructor(d) { this.#data = d; }
    #getRaw() { return this.#data; }

    compare(other) {
        return this.#getRaw() === other.#getRaw();
    }
}
const s1 = new EncryptedString("ABC");
const s2 = new EncryptedString("ABC");
console.log(s1.compare(s2));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_class_private_method_this_binding_in_callbacks() {
    let src = r#"
class TaskRunner {
    #id = 99;
    #logId() { return `Task_${this.#id}`; }

    execute() {
        const callback = () => this.#logId();
        return callback();
    }
}
console.log(new TaskRunner().execute());
"#;
    assert_eq!(run_js(src), vec!["Task_99"]);
}

#[test]
fn test_js_class_private_method_subclass_shadowing() {
    let src = r#"
class Parent {
    #action() { return "ParentPrivate"; }
    callParent() { return this.#action(); }
}
class Child extends Parent {
    #action() { return "ChildPrivate"; }
    callChild() { return this.#action(); }
}
const c = new Child();
console.log(`${c.callParent()}|${c.callChild()}`);
"#;
    assert_eq!(run_js(src), vec!["ParentPrivate|ChildPrivate"]);
}

#[test]
fn test_js_class_private_method_prototype_sharing() {
    let src = r#"
class Foo {
    #shared() { return 42; }
    getShared(other) { return other.#shared(); }
}
const f1 = new Foo();
const f2 = new Foo();
// Private methods are shared on brand check registry across instances
console.log(f1.getShared(f2));
"#;
    assert_eq!(run_js(src), vec!["42"]);
}

#[test]
fn test_js_class_private_getter_only_without_setter_throws_in_strict() {
    let src = r#"
class ConstantHolder {
    get #constVal() { return "Immutable"; }
    trySet() {
        "use strict";
        try {
            this.#constVal = "NewVal";
        } catch (e) {
            console.log("Private Getter Set TypeError");
        }
    }
}
new ConstantHolder().trySet();
"#;
    assert_eq!(run_js(src), vec!["Private Getter Set TypeError"]);
}

#[test]
fn test_js_class_private_setter_only_returns_undefined_on_get() {
    let src = r#"
class WriteOnly {
    set #secret(v) { console.log("Set Secret: " + v); }
    setSecret(v) { this.#secret = v; }
    getSecret() {
        try {
            return this.#secret;
        } catch (e) {
            return "Get Non-Existent Getter Error";
        }
    }
}
const w = new WriteOnly();
w.setSecret("Pass");
console.log(w.getSecret());
"#;
    assert_eq!(
        run_js(src),
        vec!["Set Secret: Pass", "Get Non-Existent Getter Error"]
    );
}

#[test]
fn test_js_class_private_method_unbound_this_call_throws() {
    let src = r#"
class Detached {
    #method() { return "DetachedResult"; }
    getDetached() {
        const fn = this.#method;
        return fn(); // Calling detached private method without 'this' receiver throws TypeError!
    }
}
try {
    new Detached().getDetached();
} catch (e) {
    console.log("Detached Private Call TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Detached Private Call TypeError"]);
}

#[test]
fn test_js_class_private_method_with_default_parameters() {
    let src = r#"
class Helper {
    #format(val, prefix = "[INFO]") {
        return `${prefix} ${val}`;
    }
    info(msg) { return this.#format(msg); }
}
console.log(new Helper().info("System Ready"));
"#;
    assert_eq!(run_js(src), vec!["[INFO] System Ready"]);
}

#[test]
fn test_js_class_private_method_rest_parameters() {
    let src = r#"
class Multiplier {
    #product(...nums) {
        return nums.reduce((a, b) => a * b, 1);
    }
    compute(...args) { return this.#product(...args); }
}
console.log(new Multiplier().compute(2, 3, 4));
"#;
    assert_eq!(run_js(src), vec!["24"]);
}

#[test]
fn test_js_class_private_method_recursion() {
    let src = r#"
class MathUtils {
    #factorial(n) {
        if (n <= 1) return 1;
        return n * this.#factorial(n - 1);
    }
    fact(n) { return this.#factorial(n); }
}
console.log(new MathUtils().fact(5));
"#;
    assert_eq!(run_js(src), vec!["120"]);
}

#[test]
fn test_js_class_private_method_symbol_return() {
    let src = r#"
class KeyGen {
    #genSymbol(desc) { return Symbol(desc); }
    createKey(desc) {
        const s = this.#genSymbol(desc);
        return typeof s === "symbol";
    }
}
console.log(new KeyGen().createKey("auth"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_class_private_accessor_destructuring_assignment() {
    let src = r#"
class Coords {
    #x = 10; #y = 20;
    get #pair() { return [this.#x, this.#y]; }
    set #pair([x, y]) { this.#x = x; this.#y = y; }

    update(x, y) {
        this.#pair = [x, y];
    }
    read() {
        const [x, y] = this.#pair;
        return `${x}:${y}`;
    }
}
const c = new Coords();
c.update(100, 200);
console.log(c.read());
"#;
    assert_eq!(run_js(src), vec!["100:200"]);
}

#[test]
fn test_js_class_private_method_name_inference() {
    let src = r#"
class Inspector {
    #targetMethod() {}
    getMethodName() {
        return this.#targetMethod.name;
    }
}
console.log(new Inspector().getMethodName());
"#;
    assert_eq!(run_js(src), vec!["#targetMethod"]);
}

#[test]
fn test_js_class_private_method_redefinition_throws_syntaxerror() {
    let src = r#"
try {
    eval("class Dup { #fn() {} #fn() {} }");
} catch (e) {
    console.log("Duplicate Private Method SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Duplicate Private Method SyntaxError"]);
}

#[test]
fn test_js_class_static_private_getter_setter_pair() {
    let src = r#"
class System {
    static #val = 0;
    static get #secret() { return System.#val; }
    static set #secret(v) { System.#val = v; }

    static update(v) {
        System.#secret = v;
        return System.#secret;
    }
}
console.log(System.update(42));
"#;
    assert_eq!(run_js(src), vec!["42"]);
}
