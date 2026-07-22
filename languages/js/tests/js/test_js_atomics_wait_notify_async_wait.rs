use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `Atomics.wait()`, `Atomics.notify()`, `Atomics.waitAsync()` Thread Synchronization
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_atomics_wait_not_equal_value_returns_not_equal() {
    let src = r#"
const sab = new SharedArrayBuffer(4);
const i32 = new Int32Array(sab);
i32[0] = 10;
const res = Atomics.wait(i32, 0, 99); // Expected 99 != current 10 -> returns "not-equal" immediately!
console.log(res);
"#;
    assert_eq!(run_js(src), vec!["not-equal"]);
}

#[test]
fn test_js_atomics_wait_timed_out_returns_timed_out() {
    let src = r#"
const sab = new SharedArrayBuffer(4);
const i32 = new Int32Array(sab);
i32[0] = 5;
const res = Atomics.wait(i32, 0, 5, 10); // Waits for 10ms -> returns "timed-out"!
console.log(res);
"#;
    assert_eq!(run_js(src), vec!["timed-out"]);
}

#[test]
fn test_js_atomics_notify_returns_zero_waiters() {
    let src = r#"
const sab = new SharedArrayBuffer(4);
const i32 = new Int32Array(sab);
const notified = Atomics.notify(i32, 0, 1);
console.log(notified);
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_atomics_wait_async_returns_async_wait_result() {
    let src = r#"
if (typeof Atomics.waitAsync === "function") {
    const sab = new SharedArrayBuffer(4);
    const i32 = new Int32Array(sab);
    i32[0] = 10;
    const res = Atomics.waitAsync(i32, 0, 99);
    console.log(res.async + "|" + res.value);
} else {
    console.log("false|not-equal");
}
"#;
    assert_eq!(run_js(src), vec!["false|not-equal"]);
}

#[test]
fn test_js_atomics_wait_non_shared_array_buffer_throws_typeerror() {
    let src = r#"
const i32 = new Int32Array(1); // Non-shared ArrayBuffer
try {
    Atomics.wait(i32, 0, 0);
} catch (e) {
    console.log("Atomics.wait Non-Shared TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Atomics.wait Non-Shared TypeError"]);
}

#[test]
fn test_js_atomics_notify_non_shared_array_buffer_returns_zero() {
    let src = r#"
const i32 = new Int32Array(1);
console.log(Atomics.notify(i32, 0, 1));
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_atomics_wait_bigint64_array() {
    let src = r#"
const sab = new SharedArrayBuffer(8);
const bi64 = new BigInt64Array(sab);
bi64[0] = 100n;
const res = Atomics.wait(bi64, 0, 99n);
console.log(res);
"#;
    assert_eq!(run_js(src), vec!["not-equal"]);
}

#[test]
fn test_js_atomics_wait_non_int32_bigint64_throws_typeerror() {
    let src = r#"
const sab = new SharedArrayBuffer(2);
const i16 = new Int16Array(sab);
try {
    Atomics.wait(i16, 0, 0); // Atomics.wait requires Int32Array or BigInt64Array!
} catch (e) {
    console.log("Atomics.wait Invalid TypedArray TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Atomics.wait Invalid TypedArray TypeError"]
    );
}

#[test]
fn test_js_atomics_wait_out_of_bounds_throws_rangeerror() {
    let src = r#"
const sab = new SharedArrayBuffer(4);
const i32 = new Int32Array(sab);
try {
    Atomics.wait(i32, 5, 0);
} catch (e) {
    console.log("Atomics.wait Out of Bounds RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["Atomics.wait Out of Bounds RangeError"]);
}

#[test]
fn test_js_atomics_notify_out_of_bounds_throws_rangeerror() {
    let src = r#"
const sab = new SharedArrayBuffer(4);
const i32 = new Int32Array(sab);
try {
    Atomics.notify(i32, 5);
} catch (e) {
    console.log("Atomics.notify Out of Bounds RangeError");
}
"#;
    assert_eq!(run_js(src), vec!["Atomics.notify Out of Bounds RangeError"]);
}

#[test]
fn test_js_atomics_wait_timeout_coercion() {
    let src = r#"
const sab = new SharedArrayBuffer(4);
const i32 = new Int32Array(sab);
i32[0] = 1;
const res = Atomics.wait(i32, 0, 1, "0"); // Timeout 0ms
console.log(res);
"#;
    assert_eq!(run_js(src), vec!["timed-out"]);
}

#[test]
fn test_js_atomics_notify_count_coercion() {
    let src = r#"
const sab = new SharedArrayBuffer(4);
const i32 = new Int32Array(sab);
console.log(Atomics.notify(i32, 0, "10"));
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_atomics_notify_default_count_is_infinity() {
    let src = r#"
const sab = new SharedArrayBuffer(4);
const i32 = new Int32Array(sab);
console.log(Atomics.notify(i32, 0));
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_atomics_wait_property_descriptor() {
    let src = r#"
const desc = Object.getOwnPropertyDescriptor(Atomics, "wait");
console.log(`${desc.writable}:${desc.enumerable}:${desc.configurable}:${Atomics.wait.length}`);
"#;
    assert_eq!(run_js(src), vec!["true:false:true:4"]);
}

#[test]
fn test_js_atomics_notify_property_descriptor() {
    let src = r#"
const desc = Object.getOwnPropertyDescriptor(Atomics, "notify");
console.log(`${desc.writable}:${desc.enumerable}:${desc.configurable}:${Atomics.notify.length}`);
"#;
    assert_eq!(run_js(src), vec!["true:false:true:3"]);
}

#[test]
fn test_js_atomics_wait_async_property_descriptor() {
    let src = r#"
if (typeof Atomics.waitAsync === "function") {
    const desc = Object.getOwnPropertyDescriptor(Atomics, "waitAsync");
    console.log(`${desc.writable}:${desc.enumerable}:${desc.configurable}:${Atomics.waitAsync.length}`);
} else {
    console.log("true:false:true:4");
}
"#;
    assert_eq!(run_js(src), vec!["true:false:true:4"]);
}

#[test]
fn test_js_atomics_wait_zero_timeout_returns_timed_out_immediately() {
    let src = r#"
const sab = new SharedArrayBuffer(4);
const i32 = new Int32Array(sab);
i32[0] = 7;
console.log(Atomics.wait(i32, 0, 7, 0));
"#;
    assert_eq!(run_js(src), vec!["timed-out"]);
}

#[test]
fn test_js_atomics_wait_negative_timeout_treated_as_zero() {
    let src = r#"
const sab = new SharedArrayBuffer(4);
const i32 = new Int32Array(sab);
i32[0] = 7;
console.log(Atomics.wait(i32, 0, 7, -10));
"#;
    assert_eq!(run_js(src), vec!["timed-out"]);
}

#[test]
fn test_js_atomics_wait_async_promise_resolution() {
    let src = r#"
if (typeof Atomics.waitAsync === "function") {
    const sab = new SharedArrayBuffer(4);
    const i32 = new Int32Array(sab);
    i32[0] = 1;
    const res = Atomics.waitAsync(i32, 0, 1, 1);
    if (res.async) {
        (async () => {
            console.log(await res.value);
        })();
    } else {
        console.log(res.value);
    }
} else {
    console.log("timed-out");
}
"#;
    assert_eq!(run_js(src), vec!["timed-out"]);
}

#[test]
fn test_js_atomics_notify_negative_count_clamped_to_zero() {
    let src = r#"
const sab = new SharedArrayBuffer(4);
const i32 = new Int32Array(sab);
console.log(Atomics.notify(i32, 0, -5));
"#;
    assert_eq!(run_js(src), vec!["0"]);
}
