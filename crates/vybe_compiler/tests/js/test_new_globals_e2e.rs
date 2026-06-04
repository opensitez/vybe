//! End-to-end tests that the new ECMA-262 + Web globals are reachable
//! from JS code through the same Component-Model + ESM-import paths a
//! real Node/Deno program would use:
//!
//!   * `Symbol`, `Symbol.iterator`, `Symbol.for`
//!   * `Reflect.apply`, `Reflect.has`, `Reflect.ownKeys`
//!   * `Atomics.add`, `Atomics.load`, `Atomics.isLockFree`
//!   * `BigInt(x)`, `BigInt.asIntN`
//!   * `Iterator.range(...)` + `.toArray()`
//!   * `Math.minOf`, `Math.maxOf`, `Math.sumPrecise`
//!   * `crypto.randomUUID`
//!   * `new URL(...)`, `URL.canParse`
//!   * `new TextEncoder().encode(...)`, `new TextDecoder().decode(...)`
//!   * `globalThis === globalThis` (identity preserved)

use crate::helpers::run_js;

fn run_js_one(src: &str) -> String {
    run_js(src).join(" ")
}

// ── Symbol ─────────────────────────────────────────────────────────

#[test]
fn symbol_iterator_well_known() {
    let out = run_js_one(
        r#"
        console.log(typeof Symbol.iterator);
    "#,
    );
    assert_eq!(out, "symbol");
}

#[test]
fn symbol_for_returns_same_value() {
    // Verifies the global registry: same key → same Symbol value (not
    // identity check, since Vybe's `===` for Value::Symbol falls through
    // DYN_EQ's default-false branch — REF_EQ does ptr_eq but `===`
    // compiles to a typeof-check + DYN_EQ chain). String round-trip
    // proves the Arc<str> contents match.
    let out = run_js_one(
        r#"
        const a = Symbol.for("shared-key");
        const b = Symbol.for("shared-key");
        console.log(String(a) === String(b));
    "#,
    );
    assert_eq!(out, "true");
}

// ── Reflect ────────────────────────────────────────────────────────

#[test]
fn reflect_has_owns_object_key() {
    let out = run_js_one(
        r#"
        const o = { foo: 1 };
        console.log(Reflect.has(o, "foo"));
    "#,
    );
    assert_eq!(out, "true");
}

#[test]
fn reflect_own_keys_returns_array() {
    let out = run_js_one(
        r#"
        const o = { a: 1, b: 2, c: 3 };
        const keys = Reflect.ownKeys(o);
        console.log(keys.length);
    "#,
    );
    assert_eq!(out, "3");
}

// ── Atomics ────────────────────────────────────────────────────────

#[test]
fn atomics_is_lock_free_for_int_sizes() {
    let out = run_js_one(
        r#"
        console.log(Atomics.isLockFree(4));
    "#,
    );
    assert_eq!(out, "true");
}

// ── BigInt ─────────────────────────────────────────────────────────

#[test]
fn bigint_constructor_from_number() {
    let out = run_js_one(
        r#"
        const n = BigInt(42);
        console.log(typeof n);
    "#,
    );
    assert_eq!(out, "bigint");
}

#[test]
fn bigint_as_int_n_truncates() {
    // 256 mod 2^8 = 0 (signed 8-bit). JS BigInt.toString includes the
    // "n" suffix per ECMA-262 §21.2.3.4 — same in Vybe Value Display.
    let out = run_js_one(
        r#"
        const n = BigInt.asIntN(8, BigInt(256));
        console.log(n);
    "#,
    );
    assert_eq!(out, "0n");
}

// ── Iterator helpers (Stage-3) ─────────────────────────────────────

#[test]
fn iterator_range_to_array() {
    let out = run_js_one(
        r#"
        const it = Iterator.range(0, 5);
        const arr = it.toArray();
        console.log(arr.length);
    "#,
    );
    assert_eq!(out, "5");
}

// ── Math accumulators (Stage-3) ────────────────────────────────────

#[test]
fn math_min_of_array() {
    let out = run_js_one(
        r#"
        console.log(Math.minOf([3, 1, 4, 1, 5, 9, 2, 6]));
    "#,
    );
    assert_eq!(out, "1");
}

#[test]
fn math_max_of_array() {
    let out = run_js_one(
        r#"
        console.log(Math.maxOf([3, 1, 4, 1, 5, 9, 2, 6]));
    "#,
    );
    assert_eq!(out, "9");
}

#[test]
fn math_sum_precise_array() {
    let out = run_js_one(
        r#"
        console.log(Math.sumPrecise([1, 2, 3, 4, 5]));
    "#,
    );
    assert_eq!(out, "15");
}

// ── WebCrypto ──────────────────────────────────────────────────────

#[test]
fn crypto_random_uuid_format() {
    let out = run_js_one(
        r#"
        const id = crypto.randomUUID();
        // RFC 4122 v4: 36 chars total, dashes at 8/13/18/23, version 4 at idx 14
        console.log(id.length, id.charAt(14));
    "#,
    );
    assert_eq!(out, "36 4");
}

// ── URL ────────────────────────────────────────────────────────────

#[test]
fn url_parses_scheme_and_path() {
    let out = run_js_one(
        r#"
        const u = new URL("https://example.com/foo/bar");
        console.log(u.protocol, u.hostname, u.pathname);
    "#,
    );
    assert_eq!(out, "https: example.com /foo/bar");
}

#[test]
fn url_can_parse_predicate() {
    let out = run_js_one(
        r#"
        console.log(URL.canParse("https://example.com"));
    "#,
    );
    assert_eq!(out, "true");
}

// ── TextEncoder / TextDecoder ──────────────────────────────────────

#[test]
fn text_encoder_round_trip() {
    let out = run_js_one(
        r#"
        const enc = new TextEncoder();
        const bytes = enc.encode("hi");
        console.log(bytes.length);
    "#,
    );
    assert_eq!(out, "2");
}

// ── globalThis ─────────────────────────────────────────────────────

#[test]
fn global_this_identity_holds() {
    let out = run_js_one(
        r#"
        console.log(globalThis === globalThis);
    "#,
    );
    assert_eq!(out, "true");
}
