use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Proxy.revocable & Revocable Proxy Revocation Mechanics
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_proxy_revocable_creation_returns_proxy_and_revoke() {
    let src = r#"
const { proxy, revoke } = Proxy.revocable({ a: 1 }, {});
console.log(proxy.a);
console.log(typeof revoke);
"#;
    assert_eq!(run_js(src), vec!["1", "function"]);
}

#[test]
fn test_js_proxy_revocable_get_trap_after_revoke_throws_typeerror() {
    let src = r#"
const { proxy, revoke } = Proxy.revocable({ a: 10 }, {});
revoke();
try {
    proxy.a;
} catch (e) {
    console.log("Revoked Proxy Get TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Revoked Proxy Get TypeError"]);
}

#[test]
fn test_js_proxy_revocable_set_trap_after_revoke_throws_typeerror() {
    let src = r#"
const { proxy, revoke } = Proxy.revocable({ a: 10 }, {});
revoke();
try {
    proxy.a = 20;
} catch (e) {
    console.log("Revoked Proxy Set TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Revoked Proxy Set TypeError"]);
}

#[test]
fn test_js_proxy_revocable_has_trap_after_revoke_throws_typeerror() {
    let src = r#"
const { proxy, revoke } = Proxy.revocable({ a: 10 }, {});
revoke();
try {
    "a" in proxy;
} catch (e) {
    console.log("Revoked Proxy Has TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Revoked Proxy Has TypeError"]);
}

#[test]
fn test_js_proxy_revocable_delete_property_after_revoke_throws_typeerror() {
    let src = r#"
const { proxy, revoke } = Proxy.revocable({ a: 10 }, {});
revoke();
try {
    delete proxy.a;
} catch (e) {
    console.log("Revoked Proxy Delete TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Revoked Proxy Delete TypeError"]);
}

#[test]
fn test_js_proxy_revocable_apply_trap_after_revoke_throws_typeerror() {
    let src = r#"
function fn() { return 100; }
const { proxy, revoke } = Proxy.revocable(fn, {});
revoke();
try {
    proxy();
} catch (e) {
    console.log("Revoked Proxy Apply TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Revoked Proxy Apply TypeError"]);
}

#[test]
fn test_js_proxy_revocable_construct_trap_after_revoke_throws_typeerror() {
    let src = r#"
class Item {}
const { proxy, revoke } = Proxy.revocable(Item, {});
revoke();
try {
    new proxy();
} catch (e) {
    console.log("Revoked Proxy Construct TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Revoked Proxy Construct TypeError"]);
}

#[test]
fn test_js_proxy_revocable_multiple_revoke_calls_idempotent() {
    let src = r#"
const { proxy, revoke } = Proxy.revocable({}, {});
revoke();
revoke(); // Calling revoke multiple times is safe & idempotent
console.log("Double Revoke Succeeded");
"#;
    assert_eq!(run_js(src), vec!["Double Revoke Succeeded"]);
}

#[test]
fn test_js_proxy_revocable_own_keys_after_revoke_throws_typeerror() {
    let src = r#"
const { proxy, revoke } = Proxy.revocable({ x: 1 }, {});
revoke();
try {
    Object.keys(proxy);
} catch (e) {
    console.log("Revoked Proxy OwnKeys TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Revoked Proxy OwnKeys TypeError"]);
}

#[test]
fn test_js_proxy_revocable_get_own_property_descriptor_after_revoke_throws() {
    let src = r#"
const { proxy, revoke } = Proxy.revocable({ x: 1 }, {});
revoke();
try {
    Object.getOwnPropertyDescriptor(proxy, "x");
} catch (e) {
    console.log("Revoked Proxy Descriptor TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Revoked Proxy Descriptor TypeError"]);
}

#[test]
fn test_js_proxy_revocable_get_prototype_of_after_revoke_throws() {
    let src = r#"
const { proxy, revoke } = Proxy.revocable({}, {});
revoke();
try {
    Object.getPrototypeOf(proxy);
} catch (e) {
    console.log("Revoked Proxy GetPrototypeOf TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Revoked Proxy GetPrototypeOf TypeError"]);
}

#[test]
fn test_js_proxy_revocable_set_prototype_of_after_revoke_throws() {
    let src = r#"
const { proxy, revoke } = Proxy.revocable({}, {});
revoke();
try {
    Object.setPrototypeOf(proxy, {});
} catch (e) {
    console.log("Revoked Proxy SetPrototypeOf TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Revoked Proxy SetPrototypeOf TypeError"]);
}

#[test]
fn test_js_proxy_revocable_is_extensible_after_revoke_throws() {
    let src = r#"
const { proxy, revoke } = Proxy.revocable({}, {});
revoke();
try {
    Object.isExtensible(proxy);
} catch (e) {
    console.log("Revoked Proxy IsExtensible TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Revoked Proxy IsExtensible TypeError"]);
}

#[test]
fn test_js_proxy_revocable_prevent_extensions_after_revoke_throws() {
    let src = r#"
const { proxy, revoke } = Proxy.revocable({}, {});
revoke();
try {
    Object.preventExtensions(proxy);
} catch (e) {
    console.log("Revoked Proxy PreventExtensions TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Revoked Proxy PreventExtensions TypeError"]
    );
}

#[test]
fn test_js_proxy_revocable_define_property_after_revoke_throws() {
    let src = r#"
const { proxy, revoke } = Proxy.revocable({}, {});
revoke();
try {
    Object.defineProperty(proxy, "a", { value: 1 });
} catch (e) {
    console.log("Revoked Proxy DefineProperty TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Revoked Proxy DefineProperty TypeError"]);
}

#[test]
fn test_js_proxy_revocable_temporary_access_grant() {
    let src = r#"
function withTemporaryAccess(target, fn) {
    const { proxy, revoke } = Proxy.revocable(target, {});
    try {
        return fn(proxy);
    } finally {
        revoke();
    }
}
const secretObj = { token: "XYZ-123" };
const result = withTemporaryAccess(secretObj, p => p.token);
console.log(result);
"#;
    assert_eq!(run_js(src), vec!["XYZ-123"]);
}

#[test]
fn test_js_proxy_revocable_custom_handler_traps_before_revocation() {
    let src = r#"
const { proxy, revoke } = Proxy.revocable({ count: 5 }, {
    get(t, prop) { return t[prop] * 2; }
});
console.log(proxy.count);
revoke();
console.log("Revoked Successfully");
"#;
    assert_eq!(run_js(src), vec!["10", "Revoked Successfully"]);
}

#[test]
fn test_js_proxy_revocable_returns_undefined_for_revoke_function() {
    let src = r#"
const { proxy, revoke } = Proxy.revocable({}, {});
const res = revoke();
console.log(res === undefined);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_proxy_revocable_reflect_has_after_revoke_throws() {
    let src = r#"
const { proxy, revoke } = Proxy.revocable({ a: 1 }, {});
revoke();
try {
    Reflect.has(proxy, "a");
} catch (e) {
    console.log("Reflect Has Revoked Error");
}
"#;
    assert_eq!(run_js(src), vec!["Reflect Has Revoked Error"]);
}

#[test]
fn test_js_proxy_revocable_array_target_revocation() {
    let src = r#"
const { proxy, revoke } = Proxy.revocable([10, 20, 30], {});
console.log(proxy[0]);
revoke();
try {
    proxy.push(40);
} catch (e) {
    console.log("Array Push Revoked Error");
}
"#;
    assert_eq!(run_js(src), vec!["10", "Array Push Revoked Error"]);
}
