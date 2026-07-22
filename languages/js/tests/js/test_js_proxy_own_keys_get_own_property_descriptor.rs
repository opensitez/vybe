use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Proxy Traps (ownKeys, getOwnPropertyDescriptor)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_proxy_own_keys_trap_filters_keys() {
    let src = r#"
const target = { a: 1, b: 2, _hidden: 3 };
const proxy = new Proxy(target, {
    ownKeys(t) {
        return Object.keys(t).filter(k => !k.startsWith("_"));
    }
});
console.log(Object.keys(proxy).join(","));
"#;
    assert_eq!(run_js(src), vec!["a,b"]);
}

#[test]
fn test_js_proxy_own_keys_trap_returns_symbols() {
    let src = r#"
const s1 = Symbol("s1");
const s2 = Symbol("s2");
const target = { [s1]: 10, [s2]: 20 };
const proxy = new Proxy(target, {
    ownKeys(t) {
        return [s1]; // Filter out s2
    }
});
console.log(Object.getOwnPropertySymbols(proxy).length);
"#;
    assert_eq!(run_js(src), vec!["1"]);
}

#[test]
fn test_js_proxy_get_own_property_descriptor_trap() {
    let src = r#"
const target = { val: 42 };
const proxy = new Proxy(target, {
    getOwnPropertyDescriptor(t, prop) {
        return {
            value: t[prop] * 2,
            writable: true,
            enumerable: true,
            configurable: true
        };
    }
});
const desc = Object.getOwnPropertyDescriptor(proxy, "val");
console.log(desc.value);
"#;
    assert_eq!(run_js(src), vec!["84"]);
}

#[test]
fn test_js_proxy_own_keys_duplicate_keys_throws() {
    let src = r#"
const target = {};
const proxy = new Proxy(target, {
    ownKeys(t) {
        return ["a", "a"]; // Duplicate keys not allowed in ownKeys result!
    }
});
try {
    Object.keys(proxy);
} catch (e) {
    console.log("Duplicate OwnKeys Error");
}
"#;
    assert_eq!(run_js(src), vec!["Duplicate OwnKeys Error"]);
}

#[test]
fn test_js_proxy_own_keys_non_extensible_target_must_include_all_keys() {
    let src = r#"
const target = { x: 1, y: 2 };
Object.preventExtensions(target);
const proxy = new Proxy(target, {
    ownKeys(t) {
        return ["x"]; // Missing 'y' violates invariant for non-extensible target!
    }
});
try {
    Object.keys(proxy);
} catch (e) {
    console.log("Non-Extensible OwnKeys Invariant Error");
}
"#;
    assert_eq!(run_js(src), vec!["Non-Extensible OwnKeys Invariant Error"]);
}

#[test]
fn test_js_proxy_own_keys_non_configurable_property_must_be_returned() {
    let src = r#"
const target = {};
Object.defineProperty(target, "locked", { value: 10, configurable: false });
const proxy = new Proxy(target, {
    ownKeys(t) {
        return []; // Omitting non-configurable property violates invariant!
    }
});
try {
    Object.keys(proxy);
} catch (e) {
    console.log("Non-Configurable OwnKeys Invariant Error");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Non-Configurable OwnKeys Invariant Error"]
    );
}

#[test]
fn test_js_proxy_get_own_property_descriptor_undefined_for_missing() {
    let src = r#"
const target = {};
const proxy = new Proxy(target, {
    getOwnPropertyDescriptor(t, prop) {
        return undefined;
    }
});
console.log(Object.getOwnPropertyDescriptor(proxy, "a") === undefined);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_proxy_get_own_property_descriptor_non_configurable_invariant() {
    let src = r#"
const target = {};
Object.defineProperty(target, "locked", { value: 1, configurable: false });
const proxy = new Proxy(target, {
    getOwnPropertyDescriptor(t, prop) {
        return { value: 1, configurable: true }; // Attempt to change configurable to true violates invariant!
    }
});
try {
    Object.getOwnPropertyDescriptor(proxy, "locked");
} catch (e) {
    console.log("Descriptor Invariant Error");
}
"#;
    assert_eq!(run_js(src), vec!["Descriptor Invariant Error"]);
}

#[test]
fn test_js_proxy_own_keys_for_in_loop_filtering() {
    let src = r#"
const target = { a: 1, b: 2, c: 3 };
const proxy = new Proxy(target, {
    ownKeys(t) {
        return ["a", "c"];
    }
});
const keys = [];
for (const k in proxy) {
    keys.push(k);
}
console.log(keys.join(","));
"#;
    assert_eq!(run_js(src), vec!["a,c"]);
}

#[test]
fn test_js_proxy_own_keys_reflect_own_keys_returns_strings_and_symbols() {
    let src = r#"
const sym = Symbol("s");
const target = { str: "hello", [sym]: "world" };
const proxy = new Proxy(target, {
    ownKeys(t) {
        return ["str", sym];
    }
});
const keys = Reflect.ownKeys(proxy);
console.log(keys.length + "|" + (keys[1] === sym));
"#;
    assert_eq!(run_js(src), vec!["2|true"]);
}

#[test]
fn test_js_proxy_get_own_property_descriptor_hiding_existing_property() {
    let src = r#"
const target = { secret: "data" };
const proxy = new Proxy(target, {
    getOwnPropertyDescriptor(t, prop) {
        if (prop === "secret") return undefined;
        return Reflect.getOwnPropertyDescriptor(t, prop);
    }
});
console.log(Object.getOwnPropertyDescriptor(proxy, "secret") === undefined);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_proxy_own_keys_ordering_strings_then_symbols() {
    let src = r#"
const sym = Symbol("id");
const target = {};
const proxy = new Proxy(target, {
    ownKeys(t) {
        return ["b", "a", sym, "10"];
    }
});
console.log(Reflect.ownKeys(proxy).map(k => String(k)).join(","));
"#;
    assert_eq!(run_js(src), vec!["b,a,Symbol(id),10"]);
}

#[test]
fn test_js_proxy_get_own_property_descriptor_non_extensible_target_cannot_report_new_property() {
    let src = r#"
const target = {};
Object.preventExtensions(target);
const proxy = new Proxy(target, {
    getOwnPropertyDescriptor(t, prop) {
        return { value: 100, configurable: true, enumerable: true, writable: true };
    }
});
try {
    Object.getOwnPropertyDescriptor(proxy, "newProp");
} catch (e) {
    console.log("Non-Extensible Report New Property Error");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Non-Extensible Report New Property Error"]
    );
}

#[test]
fn test_js_proxy_own_keys_virtual_properties_enumeration() {
    let src = r#"
const proxy = new Proxy({}, {
    ownKeys() {
        return ["v1", "v2", "v3"];
    },
    getOwnPropertyDescriptor() {
        return { enumerable: true, configurable: true };
    }
});
console.log(Object.keys(proxy).join(","));
"#;
    assert_eq!(run_js(src), vec!["v1,v2,v3"]);
}

#[test]
fn test_js_proxy_get_own_property_descriptor_accessor_conversion() {
    let src = r#"
const target = { count: 10 };
const proxy = new Proxy(target, {
    getOwnPropertyDescriptor(t, prop) {
        return {
            get() { return 99; },
            enumerable: true,
            configurable: true
        };
    }
});
const desc = Object.getOwnPropertyDescriptor(proxy, "count");
console.log(typeof desc.get + "|" + proxy.count);
"#;
    assert_eq!(run_js(src), vec!["function|10"]); // Target get property uses default target getter unless get trap defined
}

#[test]
fn test_js_proxy_own_keys_trap_returns_non_array_throws() {
    let src = r#"
const proxy = new Proxy({}, {
    ownKeys() {
        return "not_an_array_or_object";
    }
});
try {
    Object.keys(proxy);
} catch (e) {
    console.log("Non-List OwnKeys Error");
}
"#;
    assert_eq!(run_js(src), vec!["Non-List OwnKeys Error"]);
}

#[test]
fn test_js_proxy_get_own_property_descriptor_defaults_all_boolean_attributes() {
    let src = r#"
const target = { a: 1 };
const proxy = new Proxy(target, {
    getOwnPropertyDescriptor(t, prop) {
        return Reflect.getOwnPropertyDescriptor(t, prop);
    }
});
const desc = Object.getOwnPropertyDescriptor(proxy, "a");
console.log(desc.writable + "|" + desc.enumerable + "|" + desc.configurable);
"#;
    assert_eq!(run_js(src), vec!["true|true|true"]);
}

#[test]
fn test_js_proxy_own_keys_with_object_assign_copy() {
    let src = r#"
const target = { x: 1, y: 2 };
const proxy = new Proxy(target, {
    ownKeys(t) {
        return ["x"];
    }
});
const copy = Object.assign({}, proxy);
console.log(Object.keys(copy).join(","));
"#;
    assert_eq!(run_js(src), vec!["x"]);
}

#[test]
fn test_js_proxy_own_keys_with_spread_operator() {
    let src = r#"
const target = { a: 10, b: 20, c: 30 };
const proxy = new Proxy(target, {
    ownKeys(t) {
        return ["a", "b"];
    }
});
const spreadObj = { ...proxy };
console.log(Object.keys(spreadObj).join(","));
"#;
    assert_eq!(run_js(src), vec!["a,b"]);
}

#[test]
fn test_js_proxy_get_own_property_descriptor_passthrough() {
    let src = r#"
const target = { z: 50 };
const proxy = new Proxy(target, {});
const desc = Object.getOwnPropertyDescriptor(proxy, "z");
console.log(desc.value);
"#;
    assert_eq!(run_js(src), vec!["50"]);
}
