use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Proxy Traps (`getPrototypeOf`, `setPrototypeOf` & Prototype Invariants)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_proxy_getprototypeof_trap_basic() {
    let src = r#"
const customProto = { isCustom: true };
const proxy = new Proxy({}, {
    getPrototypeOf() {
        return customProto;
    }
});
console.log(Object.getPrototypeOf(proxy) === customProto);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_proxy_setprototypeof_trap_basic() {
    let src = r#"
const target = {};
const newProto = { a: 1 };
const proxy = new Proxy(target, {
    setPrototypeOf(t, proto) {
        console.log("setPrototypeOf trap triggered");
        return Reflect.setPrototypeOf(t, proto);
    }
});
Object.setPrototypeOf(proxy, newProto);
console.log(Object.getPrototypeOf(proxy) === newProto);
"#;
    assert_eq!(run_js(src), vec!["setPrototypeOf trap triggered", "true"]);
}

#[test]
fn test_js_proxy_getprototypeof_non_extensible_target_invariant_throws() {
    let src = r#"
const target = Object.preventExtensions({ a: 1 });
const actualProto = Object.getPrototypeOf(target);
const wrongProto = {};

const proxy = new Proxy(target, {
    getPrototypeOf() {
        return wrongProto; // Invariant: If target is non-extensible, getPrototypeOf must return target's actual prototype!
    }
});
try {
    Object.getPrototypeOf(proxy);
} catch (e) {
    console.log("getPrototypeOf Non-Extensible Invariant TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["getPrototypeOf Non-Extensible Invariant TypeError"]
    );
}

#[test]
fn test_js_proxy_setprototypeof_non_extensible_target_invariant_throws() {
    let src = r#"
const target = Object.preventExtensions({});
const proxy = new Proxy(target, {
    setPrototypeOf() {
        return true; // Returns true without changing non-extensible target prototype!
    }
});
try {
    Object.setPrototypeOf(proxy, { newProto: true });
} catch (e) {
    console.log("setPrototypeOf Non-Extensible Invariant TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["setPrototypeOf Non-Extensible Invariant TypeError"]
    );
}

#[test]
fn test_js_proxy_setprototypeof_returns_false_throws_in_strict() {
    let src = r#"
const proxy = new Proxy({}, {
    setPrototypeOf() { return false; }
});
try {
    "use strict";
    Object.setPrototypeOf(proxy, {});
} catch (e) {
    console.log("setPrototypeOf Returned False TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["setPrototypeOf Returned False TypeError"]);
}

#[test]
fn test_js_proxy_getprototypeof_non_object_non_null_return_throws() {
    let src = r#"
const proxy = new Proxy({}, {
    getPrototypeOf() { return "not_an_object"; }
});
try {
    Object.getPrototypeOf(proxy);
} catch (e) {
    console.log("getPrototypeOf Non-Object Return TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["getPrototypeOf Non-Object Return TypeError"]
    );
}

#[test]
fn test_js_proxy_getprototypeof_null_prototype() {
    let src = r#"
const proxy = new Proxy({}, {
    getPrototypeOf() { return null; }
});
console.log(Object.getPrototypeOf(proxy) === null);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_proxy_getprototypeof_instanceof_operator_trigger() {
    let src = r#"
class CustomType {}
const proxy = new Proxy({}, {
    getPrototypeOf() { return CustomType.prototype; }
});
console.log(proxy instanceof CustomType);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_proxy_getprototypeof_reflect_utility() {
    let src = r#"
const proto = { p: 10 };
const target = Object.create(proto);
const proxy = new Proxy(target, {});
console.log(Reflect.getPrototypeOf(proxy) === proto);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_proxy_setprototypeof_reflect_utility() {
    let src = r#"
const target = {};
const newProto = { p: 20 };
const proxy = new Proxy(target, {});
const success = Reflect.setPrototypeOf(proxy, newProto);
console.log(success + "|" + (Reflect.getPrototypeOf(proxy) === newProto));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_proxy_getprototypeof_trap_this_binding() {
    let src = r#"
let trapThis;
const handler = {
    getPrototypeOf(t) {
        trapThis = this;
        return Reflect.getPrototypeOf(t);
    }
};
const proxy = new Proxy({}, handler);
Object.getPrototypeOf(proxy);
console.log(trapThis === handler);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_proxy_setprototypeof_trap_this_binding() {
    let src = r#"
let trapThis;
const handler = {
    setPrototypeOf(t, proto) {
        trapThis = this;
        return Reflect.setPrototypeOf(t, proto);
    }
};
const proxy = new Proxy({}, handler);
Object.setPrototypeOf(proxy, {});
console.log(trapThis === handler);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_proxy_getprototypeof_dunder_proto_accessor() {
    let src = r#"
const customProto = { val: 100 };
const proxy = new Proxy({}, {
    getPrototypeOf() { return customProto; }
});
console.log(proxy.__proto__ === customProto);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_proxy_setprototypeof_dunder_proto_assignment() {
    let src = r#"
const target = {};
const customProto = { val: 200 };
const proxy = new Proxy(target, {
    setPrototypeOf(t, proto) {
        return Reflect.setPrototypeOf(t, proto);
    }
});
proxy.__proto__ = customProto;
console.log(Object.getPrototypeOf(proxy) === customProto);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_proxy_setprototypeof_cycle_detection_throws_typeerror() {
    let src = r#"
const target = {};
const proxy = new Proxy(target, {
    setPrototypeOf(t, proto) {
        return Reflect.setPrototypeOf(t, proto);
    }
});
try {
    Object.setPrototypeOf(target, proxy); // Cyclic prototype chain assignment throws TypeError!
} catch (e) {
    console.log("Cyclic Prototype Assignment TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Cyclic Prototype Assignment TypeError"]);
}

#[test]
fn test_js_proxy_getprototypeof_non_object_target_throws_typeerror() {
    let src = r#"
try {
    Reflect.getPrototypeOf("not_an_object");
} catch (e) {
    console.log("Reflect.getPrototypeOf Non-Object TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Reflect.getPrototypeOf Non-Object TypeError"]
    );
}

#[test]
fn test_js_proxy_setprototypeof_non_object_target_throws_typeerror() {
    let src = r#"
try {
    Reflect.setPrototypeOf(12345, {});
} catch (e) {
    console.log("Reflect.setPrototypeOf Non-Object TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Reflect.setPrototypeOf Non-Object TypeError"]
    );
}

#[test]
fn test_js_proxy_setprototypeof_null_prototype() {
    let src = r#"
const target = { a: 1 };
const proxy = new Proxy(target, {});
Object.setPrototypeOf(proxy, null);
console.log(Object.getPrototypeOf(proxy) === null);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_proxy_getprototypeof_revoked_proxy_throws() {
    let src = r#"
const { proxy, revoke } = Proxy.revocable({}, {});
revoke();
try {
    Object.getPrototypeOf(proxy);
} catch (e) {
    console.log("Revoked Proxy getPrototypeOf TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Revoked Proxy getPrototypeOf TypeError"]);
}

#[test]
fn test_js_proxy_setprototypeof_revoked_proxy_throws() {
    let src = r#"
const { proxy, revoke } = Proxy.revocable({}, {});
revoke();
try {
    Object.setPrototypeOf(proxy, {});
} catch (e) {
    console.log("Revoked Proxy setPrototypeOf TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Revoked Proxy setPrototypeOf TypeError"]);
}
