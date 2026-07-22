use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Proxy Traps (get, set, has, deleteProperty)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_proxy_get_trap_intercepts_property_access() {
    let src = r#"
const target = { a: 1 };
const proxy = new Proxy(target, {
    get(t, prop, receiver) {
        return prop in t ? t[prop] * 10 : 404;
    }
});
console.log(proxy.a + "|" + proxy.missing);
"#;
    assert_eq!(run_js(src), vec!["10|404"]);
}

#[test]
fn test_js_proxy_set_trap_validates_value_assignment() {
    let src = r#"
const target = { age: 20 };
const proxy = new Proxy(target, {
    set(t, prop, val, receiver) {
        if (prop === "age" && val < 0) {
            throw new RangeError("Age cannot be negative");
        }
        t[prop] = val;
        return true;
    }
});
proxy.age = 25;
console.log(proxy.age);
try {
    proxy.age = -5;
} catch (e) {
    console.log("RangeError Caught");
}
"#;
    assert_eq!(run_js(src), vec!["25", "RangeError Caught"]);
}

#[test]
fn test_js_proxy_has_trap_intercepts_in_operator() {
    let src = r#"
const target = { _secret: 42, public: 1 };
const proxy = new Proxy(target, {
    has(t, prop) {
        if (prop.startsWith("_")) return false;
        return prop in t;
    }
});
console.log(("_secret" in proxy) + "|" + ("public" in proxy));
"#;
    assert_eq!(run_js(src), vec!["false|true"]);
}

#[test]
fn test_js_proxy_delete_property_trap_intercepts_delete() {
    let src = r#"
const target = { a: 1, protectedKey: 2 };
const proxy = new Proxy(target, {
    deleteProperty(t, prop) {
        if (prop === "protectedKey") return false; // Delete denied
        delete t[prop];
        return true;
    }
});
console.log(delete proxy.a);
console.log(delete proxy.protectedKey);
console.log(target.protectedKey);
"#;
    assert_eq!(run_js(src), vec!["true", "false", "2"]);
}

#[test]
fn test_js_proxy_get_trap_receiver_this_identity() {
    let src = r#"
const target = { name: "Target" };
let capturedReceiver;
const proxy = new Proxy(target, {
    get(t, prop, receiver) {
        capturedReceiver = receiver;
        return t[prop];
    }
});
proxy.name;
console.log(capturedReceiver === proxy);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_proxy_set_trap_return_false_strict_mode_throws() {
    let src = r#"
const target = {};
const proxy = new Proxy(target, {
    set(t, prop, val) {
        return false; // Indicating mutation rejection
    }
});
try {
    "use strict";
    proxy.foo = "bar";
} catch (e) {
    console.log("TypeError on Set Returning False");
}
"#;
    assert_eq!(run_js(src), vec!["TypeError on Set Returning False"]);
}

#[test]
fn test_js_proxy_non_writable_non_configurable_invariant_get() {
    let src = r#"
const target = {};
Object.defineProperty(target, "fixed", {
    value: 100,
    writable: false,
    configurable: false
});
const proxy = new Proxy(target, {
    get(t, prop) {
        return 999; // Attempt to violate invariant!
    }
});
try {
    proxy.fixed;
} catch (e) {
    console.log("Proxy Invariant Get Violation");
}
"#;
    assert_eq!(run_js(src), vec!["Proxy Invariant Get Violation"]);
}

#[test]
fn test_js_proxy_non_writable_non_configurable_invariant_set() {
    let src = r#"
const target = {};
Object.defineProperty(target, "fixed", {
    value: 100,
    writable: false,
    configurable: false
});
const proxy = new Proxy(target, {
    set(t, prop, val) {
        return true;
    }
});
try {
    proxy.fixed = 999;
} catch (e) {
    console.log("Proxy Invariant Set Violation");
}
"#;
    assert_eq!(run_js(src), vec!["Proxy Invariant Set Violation"]);
}

#[test]
fn test_js_proxy_passthrough_empty_handler() {
    let src = r#"
const target = { x: 10 };
const proxy = new Proxy(target, {});
proxy.x = 20;
console.log(target.x);
console.log("x" in proxy);
delete proxy.x;
console.log(target.x);
"#;
    assert_eq!(run_js(src), vec!["20", "true", "undefined"]);
}

#[test]
fn test_js_proxy_get_trap_symbol_property() {
    let src = r#"
const sym = Symbol("test");
const target = { [sym]: "original" };
const proxy = new Proxy(target, {
    get(t, prop) {
        return typeof prop === "symbol" ? "intercepted_symbol" : t[prop];
    }
});
console.log(proxy[sym] + "|" + proxy.regular);
"#;
    assert_eq!(run_js(src), vec!["intercepted_symbol|undefined"]);
}

#[test]
fn test_js_proxy_has_trap_bypasses_prototype_chain() {
    let src = r#"
const parent = { inherited: 100 };
const child = Object.create(parent);
const proxy = new Proxy(child, {
    has(t, prop) {
        return Object.hasOwn(t, prop); // Only own properties
    }
});
console.log("inherited" in proxy);
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_proxy_delete_property_non_configurable_invariant() {
    let src = r#"
const target = {};
Object.defineProperty(target, "locked", { value: 1, configurable: false });
const proxy = new Proxy(target, {
    deleteProperty(t, prop) {
        return true; // Pretends to delete non-configurable property
    }
});
try {
    delete proxy.locked;
} catch (e) {
    console.log("Delete Non-Configurable Invariant Error");
}
"#;
    assert_eq!(run_js(src), vec!["Delete Non-Configurable Invariant Error"]);
}

#[test]
fn test_js_proxy_get_trap_method_binding_restoration() {
    let src = r#"
const obj = {
    multiplier: 3,
    calc(x) { return x * this.multiplier; }
};
const proxy = new Proxy(obj, {
    get(t, prop, receiver) {
        const val = Reflect.get(t, prop, receiver);
        return typeof val === "function" ? val.bind(receiver) : val;
    }
});
const fn = proxy.calc;
console.log(fn(5));
"#;
    assert_eq!(run_js(src), vec!["15"]);
}

#[test]
fn test_js_proxy_nested_proxy_traps() {
    let src = r#"
const target = { value: 1 };
const proxy1 = new Proxy(target, {
    get(t, prop) { return t[prop] + 10; }
});
const proxy2 = new Proxy(proxy1, {
    get(t, prop) { return t[prop] * 2; }
});
console.log(proxy2.value);
"#;
    assert_eq!(run_js(src), vec!["22"]);
}

#[test]
fn test_js_proxy_set_trap_creates_new_properties() {
    let src = r#"
const target = {};
const proxy = new Proxy(target, {
    set(t, prop, val) {
        t["prefix_" + prop] = val;
        return true;
    }
});
proxy.data = 100;
console.log(target.prefix_data);
"#;
    assert_eq!(run_js(src), vec!["100"]);
}

#[test]
fn test_js_proxy_get_trap_virtual_properties() {
    let src = r#"
const proxy = new Proxy({}, {
    get(t, prop) {
        return `Virtual_${prop}`;
    }
});
console.log(proxy.foo + "|" + proxy.bar);
"#;
    assert_eq!(run_js(src), vec!["Virtual_foo|Virtual_bar"]);
}

#[test]
fn test_js_proxy_array_index_interception() {
    let src = r#"
const arr = [10, 20, 30];
const proxy = new Proxy(arr, {
    get(t, prop) {
        if (prop === "-1") return t[t.length - 1]; // Negative index support!
        return t[prop];
    }
});
console.log(proxy["-1"]);
"#;
    assert_eq!(run_js(src), vec!["30"]);
}

#[test]
fn test_js_proxy_has_trap_non_extensible_target_invariant() {
    let src = r#"
const target = { a: 1 };
Object.preventExtensions(target);
const proxy = new Proxy(target, {
    has(t, prop) {
        return false; // Pretend 'a' doesn't exist when target is non-extensible
    }
});
try {
    "a" in proxy;
} catch (e) {
    console.log("Has Trap Non-Extensible Invariant Error");
}
"#;
    assert_eq!(run_js(src), vec!["Has Trap Non-Extensible Invariant Error"]);
}

#[test]
fn test_js_proxy_set_trap_receiver_prototype_chain() {
    let src = r#"
const proto = new Proxy({}, {
    set(t, prop, val, receiver) {
        receiver["store_" + prop] = val;
        return true;
    }
});
const child = Object.create(proto);
child.field = 50;
console.log(child.store_field);
"#;
    assert_eq!(run_js(src), vec!["50"]);
}

#[test]
fn test_js_proxy_delete_property_array_length_shrink() {
    let src = r#"
const arr = [1, 2, 3];
const proxy = new Proxy(arr, {
    deleteProperty(t, prop) {
        delete t[prop];
        return true;
    }
});
delete proxy[1];
console.log(arr.length + "|" + (1 in arr));
"#;
    assert_eq!(run_js(src), vec!["3|false"]);
}
