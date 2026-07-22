use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Proxy Traps (`getOwnPropertyDescriptor`, `defineProperty` & Invariant Validation)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_proxy_getownpropertydescriptor_trap_basic() {
    let src = r#"
const target = { a: 10 };
const proxy = new Proxy(target, {
    getOwnPropertyDescriptor(t, prop) {
        return { value: 99, writable: true, enumerable: true, configurable: true };
    }
});
const desc = Object.getOwnPropertyDescriptor(proxy, "a");
console.log(desc.value);
"#;
    assert_eq!(run_js(src), vec!["99"]);
}

#[test]
fn test_js_proxy_defineproperty_trap_basic() {
    let src = r#"
const target = {};
const proxy = new Proxy(target, {
    defineProperty(t, prop, desc) {
        console.log(`Defined: ${prop}=${desc.value}`);
        return Reflect.defineProperty(t, prop, desc);
    }
});
Object.defineProperty(proxy, "x", { value: 42, writable: true });
console.log(proxy.x);
"#;
    assert_eq!(run_js(src), vec!["Defined: x=42", "42"]);
}

#[test]
fn test_js_proxy_getownpropertydescriptor_returns_undefined_for_missing() {
    let src = r#"
const proxy = new Proxy({}, {
    getOwnPropertyDescriptor(t, prop) {
        return undefined;
    }
});
console.log(Object.getOwnPropertyDescriptor(proxy, "foo") === undefined);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_proxy_getownpropertydescriptor_non_configurable_invariant_violation_throws() {
    let src = r#"
const target = {};
Object.defineProperty(target, "fixed", { value: 1, configurable: false });
const proxy = new Proxy(target, {
    getOwnPropertyDescriptor(t, prop) {
        return { value: 1, configurable: true }; // Invariant: Cannot report non-configurable property as configurable!
    }
});
try {
    Object.getOwnPropertyDescriptor(proxy, "fixed");
} catch (e) {
    console.log("Descriptor Invariant TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Descriptor Invariant TypeError"]);
}

#[test]
fn test_js_proxy_getownpropertydescriptor_non_extensible_missing_property_invariant_throws() {
    let src = r#"
const target = Object.preventExtensions({ a: 1 });
const proxy = new Proxy(target, {
    getOwnPropertyDescriptor(t, prop) {
        return { value: 2, configurable: true, enumerable: true, writable: true }; // Invariant: Non-existent property on non-extensible target!
    }
});
try {
    Object.getOwnPropertyDescriptor(proxy, "b");
} catch (e) {
    console.log("Non-Extensible Missing Property Descriptor TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Non-Extensible Missing Property Descriptor TypeError"]
    );
}

#[test]
fn test_js_proxy_defineproperty_returning_false_throws_typeerror_in_strict() {
    let src = r#"
const proxy = new Proxy({}, {
    defineProperty() { return false; }
});
try {
    "use strict";
    Object.defineProperty(proxy, "a", { value: 1 });
} catch (e) {
    console.log("DefineProperty Trap Returned False TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["DefineProperty Trap Returned False TypeError"]
    );
}

#[test]
fn test_js_proxy_defineproperty_non_configurable_non_existent_invariant_throws() {
    let src = r#"
const target = {};
const proxy = new Proxy(target, {
    defineProperty(t, prop, desc) {
        return true; // Trap reports success without actually defining non-configurable property on target!
    }
});
try {
    Object.defineProperty(proxy, "a", { value: 1, configurable: false });
} catch (e) {
    console.log("DefineProperty Non-Configurable Invariant TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["DefineProperty Non-Configurable Invariant TypeError"]
    );
}

#[test]
fn test_js_proxy_getownpropertydescriptor_trap_symbol_property() {
    let src = r#"
const sym = Symbol("id");
const proxy = new Proxy({}, {
    getOwnPropertyDescriptor(t, prop) {
        if (prop === sym) return { value: "SymValue", configurable: true, enumerable: true };
    }
});
console.log(Object.getOwnPropertyDescriptor(proxy, sym).value);
"#;
    assert_eq!(run_js(src), vec!["SymValue"]);
}

#[test]
fn test_js_proxy_defineproperty_trap_symbol_property() {
    let src = r#"
const sym = Symbol("id");
const target = {};
const proxy = new Proxy(target, {
    defineProperty(t, prop, desc) {
        return Reflect.defineProperty(t, prop, desc);
    }
});
Object.defineProperty(proxy, sym, { value: "SymbolDefined" });
console.log(target[sym]);
"#;
    assert_eq!(run_js(src), vec!["SymbolDefined"]);
}

#[test]
fn test_js_proxy_getownpropertydescriptor_trap_receiver_this_binding() {
    let src = r#"
let trapThis;
const handler = {
    getOwnPropertyDescriptor(t, prop) {
        trapThis = this;
        return Reflect.getOwnPropertyDescriptor(t, prop);
    }
};
const proxy = new Proxy({ a: 1 }, handler);
Object.getOwnPropertyDescriptor(proxy, "a");
console.log(trapThis === handler);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_proxy_defineproperty_trap_receiver_this_binding() {
    let src = r#"
let trapThis;
const handler = {
    defineProperty(t, prop, desc) {
        trapThis = this;
        return Reflect.defineProperty(t, prop, desc);
    }
};
const proxy = new Proxy({}, handler);
Object.defineProperty(proxy, "a", { value: 1 });
console.log(trapThis === handler);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_proxy_getownpropertydescriptor_trap_non_object_return_throws() {
    let src = r#"
const proxy = new Proxy({ a: 1 }, {
    getOwnPropertyDescriptor() { return "not_an_object"; }
});
try {
    Object.getOwnPropertyDescriptor(proxy, "a");
} catch (e) {
    console.log("Descriptor Trap Non-Object TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Descriptor Trap Non-Object TypeError"]);
}

#[test]
fn test_js_proxy_defineproperty_assignment_syntax_triggers_trap() {
    let src = r#"
const target = {};
let trapCalled = false;
const proxy = new Proxy(target, {
    defineProperty(t, prop, desc) {
        trapCalled = true;
        return Reflect.defineProperty(t, prop, desc);
    }
});
proxy.newProp = 100; // Assignment on non-existent property invokes defineProperty trap!
console.log(proxy.newProp + "|TrapCalled=" + trapCalled);
"#;
    assert_eq!(run_js(src), vec!["100|TrapCalled=true"]);
}

#[test]
fn test_js_proxy_getownpropertydescriptor_reflect_forwarding() {
    let src = r#"
const target = { x: 50 };
const proxy = new Proxy(target, {
    getOwnPropertyDescriptor(t, prop) {
        const desc = Reflect.getOwnPropertyDescriptor(t, prop);
        desc.value *= 2;
        return desc;
    }
});
console.log(Object.getOwnPropertyDescriptor(proxy, "x").value);
"#;
    assert_eq!(run_js(src), vec!["100"]);
}

#[test]
fn test_js_proxy_defineproperty_read_only_target_throws() {
    let src = r#"
const target = Object.freeze({ fixed: 1 });
const proxy = new Proxy(target, {
    defineProperty(t, prop, desc) {
        return Reflect.defineProperty(t, prop, desc);
    }
});
try {
    "use strict";
    Object.defineProperty(proxy, "fixed", { value: 2 });
} catch (e) {
    console.log("DefineProperty Frozen Target TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["DefineProperty Frozen Target TypeError"]);
}

#[test]
fn test_js_proxy_getownpropertydescriptor_in_hasown_utility() {
    let src = r#"
const proxy = new Proxy({ a: 1 }, {
    getOwnPropertyDescriptor(t, prop) {
        return Reflect.getOwnPropertyDescriptor(t, prop);
    }
});
console.log(Object.hasOwn(proxy, "a") + "|" + Object.hasOwn(proxy, "b"));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_proxy_defineproperty_validation_interceptor() {
    let src = r#"
const proxy = new Proxy({}, {
    defineProperty(t, prop, desc) {
        if (typeof desc.value !== "number") throw new TypeError("Must be number");
        return Reflect.defineProperty(t, prop, desc);
    }
});
proxy.age = 25;
console.log(proxy.age);
try {
    proxy.age = "invalid";
} catch (e) {
    console.log(e.message);
}
"#;
    assert_eq!(run_js(src), vec!["25", "Must be number"]);
}

#[test]
fn test_js_proxy_getownpropertydescriptor_getter_setter_descriptor() {
    let src = r#"
const proxy = new Proxy({}, {
    getOwnPropertyDescriptor(t, prop) {
        return {
            get() { return "GetterVal"; },
            configurable: true,
            enumerable: true
        };
    }
});
const desc = Object.getOwnPropertyDescriptor(proxy, "dynamicGetter");
console.log(desc.get());
"#;
    assert_eq!(run_js(src), vec!["GetterVal"]);
}

#[test]
fn test_js_proxy_defineproperty_coerces_boolean_return() {
    let src = r#"
const proxy = new Proxy({}, {
    defineProperty(t, prop, desc) {
        Reflect.defineProperty(t, prop, desc);
        return 1; // Truthy value 1 is coerced to boolean true
    }
});
Object.defineProperty(proxy, "val", { value: 10 });
console.log(proxy.val);
"#;
    assert_eq!(run_js(src), vec!["10"]);
}

#[test]
fn test_js_proxy_getownpropertydescriptor_revoked_proxy_throws() {
    let src = r#"
const { proxy, revoke } = Proxy.revocable({ a: 1 }, {});
revoke();
try {
    Object.getOwnPropertyDescriptor(proxy, "a");
} catch (e) {
    console.log("Revoked Proxy Descriptor TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Revoked Proxy Descriptor TypeError"]);
}
