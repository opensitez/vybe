// ECMA-262 §7.1.1 ToPrimitive — bundled as `__vybe_to_primitive`.
//
// Used by the JS profile's compile_binop for `+` / `-` / `*` / `/` and
// the relational operators when an operand may be an Object. Routing
// through this JS-source polyfill (rather than a host fn) keeps the
// JS method-call protocol intact, so user `valueOf` / `toString`
// bodies see `__js_this` correctly bound when invoked.
//
// Hint: "default" | "number" | "string". `+` passes "default";
// `-` / `*` / `/` and the relational operators pass "number";
// template literals pass "string".
//
// Note: dot-syntax calls (`v.valueOf()`) go through the JS method-call
// protocol that sets `__js_this`. Bracket-syntax (`v[m]()`) takes a
// different path that bypasses the protocol — keep the dispatch
// branched on hint rather than collapsing into `v[m]()`.
function __vybe_to_primitive(v, hint) {
    if (v === null) return v;
    if (typeof v !== "object") return v;
    // Arrays: toString returns the comma-joined element list regardless
    // of hint (ECMA §23.1.3.32). The dynamic property typeof check
    // below misclassifies built-in array methods on bare arrays, so
    // short-circuit before the chain.
    if (Array.isArray(v)) {
        return v.toString();
    }
    // ECMA-262 §7.1.1 step 2.b: when the value has @@toPrimitive
    // (Symbol.toPrimitive), call it first and trust its return. The
    // walker stores the computed property under the literal source
    // string `[Symbol.toPrimitive]` (Vybe doesn't model Symbols as
    // unique keys yet — they're stringified at bind time).
    var symKey = "[Symbol.toPrimitive]";
    var sym = v[symKey];
    if (typeof sym === "function") {
        var sr = v[symKey](hint);
        if (sr === null || typeof sr !== "object") return sr;
    }
    if (hint === "string") {
        var ts1 = v.toString;
        if (typeof ts1 === "function") {
            var rs1 = v.toString();
            if (rs1 === null || typeof rs1 !== "object") return rs1;
        }
        var vof1 = v.valueOf;
        if (typeof vof1 === "function") {
            var rv1 = v.valueOf();
            if (rv1 === null || typeof rv1 !== "object") return rv1;
        }
    } else {
        var vof2 = v.valueOf;
        if (typeof vof2 === "function") {
            var rv2 = v.valueOf();
            if (rv2 === null || typeof rv2 !== "object") return rv2;
        }
        var ts2 = v.toString;
        if (typeof ts2 === "function") {
            var rs2 = v.toString();
            if (rs2 === null || typeof rs2 !== "object") return rs2;
        }
    }
    return v;
}
