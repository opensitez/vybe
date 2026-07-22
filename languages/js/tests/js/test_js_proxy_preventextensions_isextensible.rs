use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Proxy Traps (`preventExtensions`, `isExtensible` & Target Invariants)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_proxy_isextensible_trap_basic() {
    let src = r#"
const target = {};
const proxy = new Proxy(target, {
    isExtensible(t) {
        console.log("isExtensible trap called");
        return Reflect.isExtensible(t);
    }
});
console.log(Object.isExtensible(proxy));
"#;
    assert_eq!(run_js(src), vec!["isExtensible trap called", "true"]);
}

#[test]
fn test_js_proxy_preventextensions_trap_basic() {
    let src = r#"
const target = {};
const proxy = new Proxy(target, {
    preventExtensions(t) {
        console.log("preventExtensions trap called");
        return Reflect.preventExtensions(t);
    }
});
Object.preventExtensions(proxy);
console.log(Object.isExtensible(proxy));
"#;
    assert_eq!(run_js(src), vec!["preventExtensions trap called", "false"]);
}

#[test]
fn test_js_proxy_isextensible_invariant_mismatch_throws_typeerror() {
    let src = r#"
const target = {}; // Target is extensible (true)
const proxy = new Proxy(target, {
    isExtensible() {
        return false; // Trap returns false, mismatched with target! -> Throws TypeError
    }
});
try {
    Object.isExtensible(proxy);
} catch (e) {
    console.log("isExtensible Mismatch TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["isExtensible Mismatch TypeError"]);
}

#[test]
fn test_js_proxy_preventextensions_false_target_extensible_invariant_throws() {
    let src = r#"
const target = {}; // Target remains extensible
const proxy = new Proxy(target, {
    preventExtensions() {
        return true; // Trap returns true without making target non-extensible! -> Throws TypeError
    }
});
try {
    Object.preventExtensions(proxy);
} catch (e) {
    console.log("preventExtensions Target Still Extensible TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["preventExtensions Target Still Extensible TypeError"]
    );
}

#[test]
fn test_js_proxy_preventextensions_returns_false_throws_in_strict() {
    let src = r#"
const proxy = new Proxy({}, {
    preventExtensions() { return false; }
});
try {
    "use strict";
    Object.preventExtensions(proxy);
} catch (e) {
    console.log("preventExtensions Returned False TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["preventExtensions Returned False TypeError"]
    );
}

#[test]
fn test_js_proxy_isextensible_reflect_utility() {
    let src = r#"
const target = { a: 1 };
const proxy = new Proxy(target, {});
console.log(Reflect.isExtensible(proxy));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_proxy_preventextensions_reflect_utility() {
    let src = r#"
const target = { a: 1 };
const proxy = new Proxy(target, {});
const success = Reflect.preventExtensions(proxy);
console.log(success + "|" + Reflect.isExtensible(proxy));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_proxy_preventextensions_blocks_property_addition() {
    let src = r#"
const proxy = new Proxy({}, {
    preventExtensions(t) {
        return Reflect.preventExtensions(t);
    }
});
Object.preventExtensions(proxy);
try {
    "use strict";
    proxy.newProp = 100;
} catch (e) {
    console.log("Add Property To Non-Extensible Proxy TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Add Property To Non-Extensible Proxy TypeError"]
    );
}

#[test]
fn test_js_proxy_isextensible_non_object_target_throws_typeerror() {
    let src = r#"
try {
    Reflect.isExtensible("not_an_object");
} catch (e) {
    console.log("Reflect.isExtensible Non-Object TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Reflect.isExtensible Non-Object TypeError"]
    );
}

#[test]
fn test_js_proxy_preventextensions_non_object_target_throws_typeerror() {
    let src = r#"
try {
    Reflect.preventExtensions(42);
} catch (e) {
    console.log("Reflect.preventExtensions Non-Object TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Reflect.preventExtensions Non-Object TypeError"]
    );
}

#[test]
fn test_js_proxy_isextensible_trap_this_binding() {
    let src = r#"
let trapThis;
const handler = {
    isExtensible(t) {
        trapThis = this;
        return Reflect.isExtensible(t);
    }
};
const proxy = new Proxy({}, handler);
Object.isExtensible(proxy);
console.log(trapThis === handler);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_proxy_preventextensions_trap_this_binding() {
    let src = r#"
let trapThis;
const handler = {
    preventExtensions(t) {
        trapThis = this;
        return Reflect.preventExtensions(t);
    }
};
const proxy = new Proxy({}, handler);
Object.preventExtensions(proxy);
console.log(trapThis === handler);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_proxy_seal_invokes_preventextensions_trap() {
    let src = r#"
let preventCalled = false;
const proxy = new Proxy({ a: 1 }, {
    preventExtensions(t) {
        preventCalled = true;
        return Reflect.preventExtensions(t);
    }
});
Object.seal(proxy);
console.log(preventCalled + "|isSealed=" + Object.isSealed(proxy));
"#;
    assert_eq!(run_js(src), vec!["true|isSealed=true"]);
}

#[test]
fn test_js_proxy_freeze_invokes_preventextensions_trap() {
    let src = r#"
let preventCalled = false;
const proxy = new Proxy({ a: 1 }, {
    preventExtensions(t) {
        preventCalled = true;
        return Reflect.preventExtensions(t);
    }
});
Object.freeze(proxy);
console.log(preventCalled + "|isFrozen=" + Object.isFrozen(proxy));
"#;
    assert_eq!(run_js(src), vec!["true|isFrozen=true"]);
}

#[test]
fn test_js_proxy_isextensible_after_preventextensions() {
    let src = r#"
const target = {};
const proxy = new Proxy(target, {});
console.log(Object.isExtensible(proxy));
Object.preventExtensions(proxy);
console.log(Object.isExtensible(proxy));
"#;
    assert_eq!(run_js(src), vec!["true", "false"]);
}

#[test]
fn test_js_proxy_isextensible_coerces_boolean_return() {
    let src = r#"
const target = {};
const proxy = new Proxy(target, {
    isExtensible() {
        return 1; // Truthy 1 matches target's true extensibility
    }
});
console.log(Object.isExtensible(proxy));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_proxy_preventextensions_already_non_extensible_target() {
    let src = r#"
const target = Object.preventExtensions({});
const proxy = new Proxy(target, {
    preventExtensions(t) {
        return Reflect.preventExtensions(t);
    }
});
console.log(Object.preventExtensions(proxy) === proxy);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_proxy_isextensible_revoked_proxy_throws() {
    let src = r#"
const { proxy, revoke } = Proxy.revocable({}, {});
revoke();
try {
    Object.isExtensible(proxy);
} catch (e) {
    console.log("Revoked Proxy isExtensible TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Revoked Proxy isExtensible TypeError"]);
}

#[test]
fn test_js_proxy_preventextensions_revoked_proxy_throws() {
    let src = r#"
const { proxy, revoke } = Proxy.revocable({}, {});
revoke();
try {
    Object.preventExtensions(proxy);
} catch (e) {
    console.log("Revoked Proxy preventExtensions TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Revoked Proxy preventExtensions TypeError"]
    );
}

#[test]
fn test_js_proxy_isextensible_on_array_target() {
    let src = r#"
const arr = [1, 2];
const proxy = new Proxy(arr, {
    isExtensible(t) { return Reflect.isExtensible(t); }
});
console.log(Object.isExtensible(proxy));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}
