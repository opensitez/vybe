# `wasm:js-*` Builtin Conventions

This file pins the contract every `wasm:js-*` import in Vybe must
follow. It covers:

1. **Type marshaling** — how WASM-level types map to JS-level values
2. **Error handling** — what a host handler does when things go wrong
3. **Null / undefined semantics**
4. **Identity preservation**
5. **Forward-compat with the CG `js-string-builtins` family**

The goal: every import in the Vybe stdlib family
(`wasm:js-array`, `wasm:js-object`, `wasm:js-map`, `wasm:js-set`,
`wasm:js-weakmap`, `wasm:js-weakset`, `wasm:js-arraybuffer`,
`wasm:js-sharedarraybuffer`, `wasm:js-dataview`,
`wasm:js-typedarray-*`, and existing `wasm:js-string`,
`wasm:js-number`, `wasm:js-boolean`, `wasm:js-undefined`,
`wasm:js-symbol`, `wasm:js-bigint`) obeys these rules uniformly.
When v8 / SpiderMonkey / JavaScriptCore implement any of these
(either today for string/primitives, or hypothetically in the future
for collections), their signatures match — so browser JS glue is a
one-liner per method, and future CG standardization is a rename.

## 1. Type marshaling

### Collection instances

Any reference to a collection (Array, Map, Set, Object, WeakMap,
WeakSet, ArrayBuffer, SharedArrayBuffer, DataView, every typed-array
variant):

```
WASM type: externref
```

The host wraps its native representation (Rust `Arc<Mutex<Object>>`,
v8's hidden-class-allocated JS object, etc.) and the user never sees
the backing structure. Passing the same externref through multiple
import calls MUST yield the same backing object — no copy-on-pass,
no identity shuffling. See §4.

### Indices, sizes, offsets, lane indices

```
WASM type: i32
```

Unsigned interpretation in all cases (negative values are either
invalid input → see §2, or the language mapping layer's concern,
not the import's). `Array.prototype.at(-1)` is a JS-surface
feature, not an import-level one — our `wasm:js-array.at` takes a
positive i32; JS `at()` bridges to it by adding the negative index
to length at the language layer.

### Booleans

```
WASM type: i32 with { 0 = false, non-zero = true }
```

Includes `littleEndian` flags on `DataView`, `deep` flags on clone
ops, `descriptor.writable` booleans, etc. Zero is always exactly
false; any non-zero i32 is true.

### Numeric values (non-typed-array)

Spec numeric coercion semantics apply. Float values use `f64`
unboxed, not a wrapper externref:

```
JS Number → WASM type: f64
```

This matches `wasm:js-number.toF64` and `.fromF64` already. Integer
numeric arguments follow:

```
Signed 32-bit context → WASM type: i32
Unsigned 32-bit context → WASM type: i32 (caller interprets)
64-bit BigInt context → WASM type: i64
```

### Typed-array element values

Per-variant, matching the typed-array element type:

| Variant | WASM type for value args/results |
|---|---|
| `Int8Array`, `Uint8Array`, `Uint8ClampedArray`, `Int16Array`, `Uint16Array`, `Int32Array`, `Uint32Array` | `i32` (sign/zero extension per variant) |
| `Float32Array` | `f32` |
| `Float64Array` | `f64` |
| `BigInt64Array`, `BigUint64Array` | `i64` |

### Strings

```
WASM type: externref
```

Follows existing `wasm:js-string-builtins` convention. A string is
an opaque externref; operations on it go through `wasm:js-string.*`
imports.

### JS values of unknown shape

Anything that might be string / number / object / null / undefined /
symbol / bigint flowing through generic positions:

```
WASM type: externref
```

This is the "bag-of-values" pattern. Our universal value ABI is
externref; all our stdlib ops accept and return externref for
generic positions.

### Multi-value results

When an import needs to return more than one value, use the
multi-value proposal (`result_arity > 1`). Result ordering:

```
(primary_result, status_code?)
```

The primary result comes first. A status code (i32) follows only
when the op can fail recoverably and the caller needs to
distinguish. Example:

```
wasm:js-map.delete(map, key)
  → (i32 was_present, externref previous_value)
```

Single-value results: no status code, just the value.

## 2. Error handling

Three error classes. Every handler must pick one per possible
failure mode — document inline.

### Class 1 — Invariant violations

**What**: OOM, internal VM bugs, impossible states (e.g. externref
that doesn't wrap a collection when one is expected).

**Behavior**: WASM trap. The engine aborts this module instance.
No recovery possible; these are "the host is broken" conditions.

```rust
// Example handler fragment
let Some(arr) = extract_array(&args[0]) else {
    return Err(VMError::trap("wasm:js-array.push called on non-array"));
};
```

### Class 2 — JS-spec errors

**What**: conditions where ECMA-262 mandates throwing a specific
error type. `TypeError` when pushing to a frozen Array. `RangeError`
on length overflow. `TypeError` when using a primitive as a
WeakMap key.

**Behavior**: throw via the exception-handling proposal. Payload is
a standard error object — `{ name: string, message: string,
stack: string }` — with:
- `name`: matches the JS constructor name exactly (`"TypeError"`,
  `"RangeError"`, `"SyntaxError"`, …)
- `message`: matches ECMA-262's canonical message when the spec
  provides one; otherwise a descriptive message that clearly
  identifies the failure
- `stack`: filled in from the current call stack, same format as
  `Error.prototype.stack`

```rust
// Example handler fragment
if array_is_frozen(&arr) {
    return vm.throw_type_error("Cannot add property to frozen array");
}
```

### Class 3 — Input errors the spec explicitly permits

**What**: cases where ECMA-262 says "return undefined" or similar
non-error behavior for bad input.

Examples:
- `Array.prototype.at(999)` on a 3-element array → returns `undefined`
- `Map.prototype.get(missing_key)` → returns `undefined`
- `Array.prototype.indexOf(notPresent)` → returns `-1`

**Behavior**: return the spec's expected result. No trap, no throw.
The language mapping layer translates `undefined` → `KeyError`
(Python) / `nil` (Ruby) / `None` if the source language requires it.

## 3. Null and undefined

JS distinguishes `null` and `undefined`. Our value ABI mirrors this:

| JS value | WASM-side externref | VM Value | Host fast-path |
|---|---|---|---|
| `null` | `ref.null externref` | `Value::Null` | JS `null` passed natively |
| `undefined` | `global.get $js_undefined` (imported from `wasm:js-undefined`) | `Value::Undefined` | JS `undefined` passed natively |

When a spec operation says "returns undefined", the handler returns
the `$js_undefined` singleton, NOT `ref.null externref`. When a spec
operation says "returns null" it returns `ref.null externref`.

## 4. Identity preservation

**Invariant**: passing the same externref into a handler and later
retrieving an externref that the spec says is "the same object"
must yield the same WASM-level identity.

Concretely:

```
let arr = wasm:js-array.new();
wasm:js-array.push(arr, v);
let elem = wasm:js-array.at(arr, 0);

// If v was an externref wrapping an Object, `elem` must be an
// externref wrapping the SAME Object (same Arc<Mutex<Object>> on
// Vybe VM; same v8-hidden-class-object on v8).
```

No defensive copies. No "let me allocate a new wrapper." The host's
job is to be a transparent carrier of JS value identity.

`Object.is(a, b)` / `===` / `Array.prototype.indexOf` (which uses
SameValueZero) all rely on this. Breaking it means these methods
misbehave.

On Vybe VM: externref is a direct `Value` clone, which for
`Value::Object(Arc<_>)` is pointer-equal after `Arc::clone`. On v8:
externref is opaque and engine-managed; identity is preserved by
construction.

## 5. Forward compatibility with the CG

The import module names use the `wasm:js-*` convention. When the
WebAssembly CG standardizes additional builtins (e.g. a hypothetical
`wasm:js-array-builtins` proposal analogous to today's
`wasm:js-string-builtins`), the migration is a one-line rename in
`crates/vybe_bytecode/src/wasm/wasm_js_array.rs` —
`pub const MODULE: &str = "wasm:js-array"` → whatever the final
standard name becomes (e.g. `"wasm:js-array-builtins"`). User
modules that were emitted against our pre-standard name can be
re-emitted with the new name; the behavioral contract stays
identical.

Signatures are designed to match what v8 / SpiderMonkey / JSCore
would expose if they implemented these natively — so browser JS
glue (see `tools/vybe-loader/vybe_js_collections.js`, Phase C) is
literally a one-to-one method wrapping:

```js
{
  "wasm:js-array": {
    "new": () => new Array(),
    "push": (arr, v) => arr.push(v),
    "pop": (arr) => arr.pop(),
    "at": (arr, i) => arr.at(i),
    // ... every other method
  }
}
```

Each line is a direct call to the native JS method. Zero
adaptation.

## 6. Capabilities

All `wasm:js-*` imports declared in the Vybe stdlib family require
the `Capabilities::Safe` level — no I/O, no threads beyond what our
VM already supports (see `wasm:js-sharedarraybuffer` for the one
exception, which gates on `Capabilities::Threads`).

## 7. Versioning

No versioning in module names (e.g. no `wasm:js-array/v1`). The
import surface evolves additively — new methods are added, existing
methods never change signature without a new method name. This
matches the CG's approach to `wasm:js-string-builtins`.
