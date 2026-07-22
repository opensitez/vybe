use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Class Static Initialization Blocks (ES2022)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_class_static_initialization_block_basic() {
    let src = r#"
class App {
    static config;
    static {
        this.config = "InitializedConfig";
    }
}
console.log(App.config);
"#;
    assert_eq!(run_js(src), vec!["InitializedConfig"]);
}

#[test]
fn test_js_class_static_block_private_field_access() {
    let src = r#"
let getPrivateVal;
class SecretContainer {
    #secret = "TopSecret";
    static {
        getPrivateVal = (instance) => instance.#secret;
    }
}
const sc = new SecretContainer();
console.log(getPrivateVal(sc));
"#;
    assert_eq!(run_js(src), vec!["TopSecret"]);
}

#[test]
fn test_js_class_static_block_execution_order_multiple_blocks() {
    let src = r#"
const log = [];
class OrderTest {
    static field1 = (() => { log.push("Field 1"); return 1; })();
    static {
        log.push("Static Block 1");
    }
    static field2 = (() => { log.push("Field 2"); return 2; })();
    static {
        log.push("Static Block 2");
    }
}
console.log(log.join("->"));
"#;
    assert_eq!(
        run_js(src),
        vec!["Field 1->Static Block 1->Field 2->Static Block 2"]
    );
}

#[test]
fn test_js_class_static_block_this_refers_to_class_constructor() {
    let src = r#"
class Target {
    static isSelf;
    static {
        this.isSelf = (this === Target);
    }
}
console.log(Target.isSelf);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_class_static_block_subclass_execution_order() {
    let src = r#"
const log = [];
class Base {
    static { log.push("Base Block"); }
}
class Sub extends Base {
    static { log.push("Sub Block"); }
}
console.log(log.join("->"));
"#;
    assert_eq!(run_js(src), vec!["Base Block->Sub Block"]);
}

#[test]
fn test_js_class_static_block_try_catch_error_handling() {
    let src = r#"
class SafeInit {
    static status;
    static {
        try {
            throw new Error("Init Error");
        } catch (e) {
            this.status = "FallbackStatus";
        }
    }
}
console.log(SafeInit.status);
"#;
    assert_eq!(run_js(src), vec!["FallbackStatus"]);
}

#[test]
fn test_js_class_static_block_var_hoisting_scoped_to_block() {
    let src = r#"
class Scoping {
    static {
        var blockVar = "VarInBlock";
    }
    static getVar() {
        return typeof blockVar === "undefined";
    }
}
console.log(Scoping.getVar());
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_class_static_block_lexical_environment_isolation() {
    let src = r#"
class Iso1 {
    static { const x = 10; }
}
class Iso2 {
    static { const x = 20; }
}
console.log("Isolated Blocks Succeeded");
"#;
    assert_eq!(run_js(src), vec!["Isolated Blocks Succeeded"]);
}

#[test]
fn test_js_class_static_block_outer_scope_access() {
    let src = r#"
const externalValue = "GlobalContext";
class Bridge {
    static val;
    static {
        this.val = externalValue.toUpperCase();
    }
}
console.log(Bridge.val);
"#;
    assert_eq!(run_js(src), vec!["GLOBALCONTEXT"]);
}

#[test]
fn test_js_class_static_block_static_private_method_export() {
    let src = r#"
let callStaticPrivate;
class Exposer {
    static #privateStaticMethod() {
        return "ExposedStaticPrivate";
    }
    static {
        callStaticPrivate = () => Exposer.#privateStaticMethod();
    }
}
console.log(callStaticPrivate());
"#;
    assert_eq!(run_js(src), vec!["ExposedStaticPrivate"]);
}

#[test]
fn test_js_class_static_block_uncaught_error_halts_class_evaluation() {
    let src = r#"
try {
    eval(`
        class Unsafe {
            static { throw new Error("ClassEvalError"); }
        }
    `);
} catch (e) {
    console.log(e.message);
}
"#;
    assert_eq!(run_js(src), vec!["ClassEvalError"]);
}

#[test]
fn test_js_class_static_block_returns_value_ignored() {
    let src = r#"
class RetTest {
    static {
        // Return statements in static blocks syntax error per spec
    }
}
console.log("No Return Static Block Succeeded");
"#;
    assert_eq!(run_js(src), vec!["No Return Static Block Succeeded"]);
}

#[test]
fn test_js_class_static_block_super_property_access() {
    let src = r#"
class Base {
    static baseValue = "BaseVal";
}
class Sub extends Base {
    static subValue;
    static {
        this.subValue = super.baseValue + "_Extended";
    }
}
console.log(Sub.subValue);
"#;
    assert_eq!(run_js(src), vec!["BaseVal_Extended"]);
}

#[test]
fn test_js_class_static_block_super_method_call() {
    let src = r#"
class Base {
    static getTitle() { return "BaseTitle"; }
}
class Sub extends Base {
    static title;
    static {
        this.title = super.getTitle().toUpperCase();
    }
}
console.log(Sub.title);
"#;
    assert_eq!(run_js(src), vec!["BASETITLE"]);
}

#[test]
fn test_js_class_static_block_await_expression_syntaxerror() {
    let src = r#"
try {
    eval("class Bad { static { await Promise.resolve(); } }");
} catch (e) {
    console.log("Await in Static Block Error");
}
"#;
    assert_eq!(run_js(src), vec!["Await in Static Block Error"]);
}

#[test]
fn test_js_class_static_block_for_loop_initialization() {
    let src = r#"
class MathTable {
    static table = [];
    static {
        for (let i = 1; i <= 3; i++) {
            this.table.push(i * 10);
        }
    }
}
console.log(MathTable.table.join(","));
"#;
    assert_eq!(run_js(src), vec!["10,20,30"]);
}

#[test]
fn test_js_class_static_block_nested_class_definition() {
    let src = r#"
class Outer {
    static Inner;
    static {
        this.Inner = class {
            static name = "InnerClass";
        };
    }
}
console.log(Outer.Inner.name);
"#;
    assert_eq!(run_js(src), vec!["InnerClass"]);
}

#[test]
fn test_js_class_static_block_anonymous_class_expression() {
    let src = r#"
const Anon = class {
    static created;
    static {
        this.created = true;
    }
};
console.log(Anon.created);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_class_static_block_destructuring_assignment() {
    let src = r#"
const initData = { env: "prod", port: 443 };
class ServerConfig {
    static env; static port;
    static {
        ({ env: this.env, port: this.port } = initData);
    }
}
console.log(`${ServerConfig.env}:${ServerConfig.port}`);
"#;
    assert_eq!(run_js(src), vec!["prod:443"]);
}

#[test]
fn test_js_class_static_block_symbol_property_initialization() {
    let src = r#"
const symKey = Symbol("registry");
class Registry {
    static {
        this[symKey] = "RegisteredSymbolValue";
    }
    static getVal() { return this[symKey]; }
}
console.log(Registry.getVal());
"#;
    assert_eq!(run_js(src), vec!["RegisteredSymbolValue"]);
}
