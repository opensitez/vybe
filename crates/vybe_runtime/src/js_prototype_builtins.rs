//! # `wasm:js-prototypes`
//!
//! Custom Descriptors — "Declarative Prototype Initialization" / "Configuration
//! API" (`proposals/custom-descriptors/.../Overview.md` §660-943).
//!
//! One function, `configureAll`, called from a module's start function to
//! populate imported prototype objects without per-method JS glue. The
//! proposal's stated motivation is startup latency: toolchains expect to
//! configure thousands of prototypes and tens of thousands of methods, and
//! doing that from JS glue is what this exists to avoid.
//!
//! ```wasm
//! (type $configureAll (func (param (ref null $prototypes))
//!                           (param (ref null $functions))
//!                           (param (ref null $data))
//!                           (param externref)))
//! ```
//!
//! The three arrays are consumed IN ORDER as the byte stream is parsed:
//! each `protoconfig` takes the next prototype, each `constructorconfig` or
//! `methodconfig` takes the next function.

use crate::value::ObjectKind;
use crate::{HostContext, VM, Value};
use std::sync::Arc;

fn trap(ctx: &mut HostContext, msg: &str) {
    ctx.throw_value(Value::String(Arc::from(msg)));
}

/// Elements of a GC array (`ObjectKind::Array`), or `None` for a null/absent
/// argument — the params are `(ref null …)`, so null is legal and means "no
/// entries", not an error.
fn array_elems(v: Option<&Value>) -> Option<Vec<Value>> {
    match v {
        Some(Value::Object(o)) => {
            let ob = o.lock().unwrap();
            match &ob.kind {
                ObjectKind::Array(a) => Some(a.clone()),
                _ => None,
            }
        }
        _ => None,
    }
}

/// The `data` array as bytes. Elements are `i8`, so they arrive as integers.
fn data_bytes(v: Option<&Value>) -> Vec<u8> {
    array_elems(v)
        .unwrap_or_default()
        .iter()
        .map(|e| match e {
            Value::I32(n) => *n as u8,
            Value::I64(n) => *n as u8,
            Value::F64(n) => *n as u8,
            _ => 0,
        })
        .collect()
}

/// A cursor over the configuration byte stream. Every read is fallible: the
/// stream is module-supplied data and a malformed one must trap, never index
/// past the end.
struct Cursor<'a> {
    b: &'a [u8],
    i: usize,
}

impl<'a> Cursor<'a> {
    fn byte(&mut self) -> Option<u8> {
        let v = *self.b.get(self.i)?;
        self.i += 1;
        Some(v)
    }
    /// Unsigned LEB128 — the `vec` length prefix and `name` byte count.
    fn u32(&mut self) -> Option<u32> {
        let mut out: u32 = 0;
        let mut shift = 0;
        loop {
            let b = self.byte()?;
            out |= ((b & 0x7F) as u32).checked_shl(shift)?;
            if b & 0x80 == 0 {
                return Some(out);
            }
            shift += 7;
            if shift > 31 {
                return None;
            }
        }
    }
    /// Signed LEB128 — `parentidx`, which is −1 for "no parent".
    fn s32(&mut self) -> Option<i32> {
        let mut out: i32 = 0;
        let mut shift = 0;
        loop {
            let b = self.byte()?;
            out |= ((b & 0x7F) as i32) << shift;
            shift += 7;
            if b & 0x80 == 0 {
                if shift < 32 && (b & 0x40) != 0 {
                    out |= -1i32 << shift;
                }
                return Some(out);
            }
            if shift > 31 {
                return None;
            }
        }
    }
    /// `name` — a WASM name: a byte count followed by UTF-8.
    fn name(&mut self) -> Option<String> {
        let n = self.u32()? as usize;
        let end = self.i.checked_add(n)?;
        let s = self.b.get(self.i..end)?;
        self.i = end;
        String::from_utf8(s.to_vec()).ok()
    }
}

/// Install one `methodconfig` kind on `target`.
///
/// Getters and setters use the runtime's own accessor spelling (`__get_<name>`
/// / `__set_<name>`) — the same keys `Object.defineProperty` writes, so a
/// configured accessor is indistinguishable from a JS-defined one.
fn install(target: &Value, kind: u8, name: &str, f: Value) -> bool {
    let Value::Object(o) = target else {
        return false;
    };
    let key = match kind {
        0x00 => name.to_string(),
        0x01 => format!("__get_{name}"),
        0x02 => format!("__set_{name}"),
        _ => return false,
    };
    o.lock().unwrap().properties.insert(key, f);
    true
}

pub fn register(vm: &mut VM) {
    vm.register_host_fn(
        "wasm:js-prototypes",
        "configureAll",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let protos = array_elems(args.first()).unwrap_or_default();
            let funcs = array_elems(args.get(1)).unwrap_or_default();
            let data = data_bytes(args.get(2));
            let ctors_obj = args.get(3).cloned().unwrap_or(Value::Undefined);

            let mut c = Cursor { b: &data, i: 0 };
            let mut pi = 0usize; // next prototype
            let mut fi = 0usize; // next function

            // ⛔ Errors are detected LAZILY and the spec says so: "user-visible
            // partial modifications may have occurred before an error is
            // thrown". So this installs as it parses rather than validating
            // the whole stream first — a pre-pass would be a nicer API and a
            // WRONG one.
            macro_rules! bail {
                ($m:expr) => {{
                    trap(ctx, &format!("{} at data offset {}", $m, c.i));
                    return Value::Undefined;
                }};
            }
            macro_rules! next_fn {
                () => {{
                    let Some(f) = funcs.get(fi).cloned() else {
                        bail!("wasm:js-prototypes: functions array exhausted")
                    };
                    fi += 1;
                    f
                }};
            }

            let Some(nprotos) = c.u32() else {
                bail!("wasm:js-prototypes: truncated protoconfig count")
            };
            for _ in 0..nprotos {
                let Some(proto) = protos.get(pi).cloned() else {
                    bail!("wasm:js-prototypes: prototypes array exhausted")
                };
                let this_proto_idx = pi;
                pi += 1;

                // vec(constructorconfig), size <= 1.
                let Some(nctor) = c.u32() else {
                    bail!("wasm:js-prototypes: truncated constructor count")
                };
                if nctor > 1 {
                    bail!("wasm:js-prototypes: at most one constructorconfig per prototype")
                }
                for _ in 0..nctor {
                    let Some(cname) = c.name() else {
                        bail!("wasm:js-prototypes: truncated constructor name")
                    };
                    let ctor = next_fn!();
                    // The constructor is installed BOTH as the prototype's
                    // `constructor` property and, by name, on the constructors
                    // object — the proposal's fourth parameter exists because
                    // constructors "cannot be added to the module's exports
                    // object".
                    if let Value::Object(p) = &proto {
                        p.lock()
                            .unwrap()
                            .properties
                            .insert("constructor".into(), ctor.clone());
                    }
                    if let Value::Object(co) = &ctors_obj {
                        co.lock().unwrap().properties.insert(cname, ctor.clone());
                    }
                    // Statics: "installed on the current constructor … not
                    // wrapped or modified in any way, and in particular it
                    // does not receive a method receiver as its first
                    // parameter."
                    let Some(nstatics) = c.u32() else {
                        bail!("wasm:js-prototypes: truncated static method count")
                    };
                    for _ in 0..nstatics {
                        let Some(kind) = c.byte() else {
                            bail!("wasm:js-prototypes: truncated static method kind")
                        };
                        let Some(mname) = c.name() else {
                            bail!("wasm:js-prototypes: truncated static method name")
                        };
                        let f = next_fn!();
                        if !install(&ctor, kind, &mname, f) {
                            bail!("wasm:js-prototypes: bad static methodconfig")
                        }
                    }
                }

                // Top-level methods go on the PROTOTYPE.
                let Some(nmethods) = c.u32() else {
                    bail!("wasm:js-prototypes: truncated method count")
                };
                for _ in 0..nmethods {
                    let Some(kind) = c.byte() else {
                        bail!("wasm:js-prototypes: truncated method kind")
                    };
                    let Some(mname) = c.name() else {
                        bail!("wasm:js-prototypes: truncated method name")
                    };
                    let f = next_fn!();
                    if !install(&proto, kind, &mname, f) {
                        bail!("wasm:js-prototypes: bad methodconfig")
                    }
                }

                // parentidx: −1 for none, else an index into the prototypes
                // array. ⛔ "The index must be less than the current prototype
                // index" — a forward or self reference is rejected, which is
                // what keeps the chain acyclic.
                let Some(parent) = c.s32() else {
                    bail!("wasm:js-prototypes: truncated parentidx")
                };
                if parent >= 0 {
                    let p = parent as usize;
                    if p >= this_proto_idx {
                        bail!("wasm:js-prototypes: parentidx must precede the current prototype")
                    }
                    let Some(parent_proto) = protos.get(p).cloned() else {
                        bail!("wasm:js-prototypes: parentidx out of range")
                    };
                    // "must be a valid prototype, i.e. a JS object or null".
                    if !matches!(parent_proto, Value::Object(_) | Value::Null) {
                        bail!("wasm:js-prototypes: parent is not a valid prototype")
                    }
                    if let Value::Object(p) = &proto {
                        p.lock()
                            .unwrap()
                            .properties
                            .insert("__proto__".into(), parent_proto);
                    }
                }
            }
            Value::Undefined
        }),
    );
}
