//! `CanonicalABI.md` §Flat Lifting / §Flat Lowering — values in and out of a
//! flattening.
//!
//! [`canon_flat`] answers what core types a component type OCCUPIES;
//! this module moves values through them. The pair is what `canon lift` and
//! `canon lower` are built from, and the two directions have to agree
//! exactly — a variant lowered with one padding rule and lifted with another
//! reads a neighbouring field.
//!
//! # Why this is not `canon_value`
//!
//! [`canon_value`] moves a value to and from LINEAR MEMORY at its canonical
//! layout. This module moves it to and from a list of CORE VALUES (registers).
//! Strings and lists appear in both: their bytes always live in memory, and
//! only the (ptr, length) pair travels flat. That overlap is deliberate in the
//! spec — "lowering can reuse the previous definitions; only the resulting
//! pointers are returned differently".
//!
//! # ⚠The coercion tables are the whole difficulty
//!
//! `flatten_variant` gives every case the SAME static core types by `join`ing
//! them, so a case whose payload is `f32` may have to travel in an `i32` slot,
//! and an `i32` in an `i64` slot. Both directions must undo that identically:
//!
//! | lower (`have` → `want`) | lift (`have` → `want`) |
//! | --- | --- |
//! | `f32` → `i32`: reinterpret bits | `i32` → `f32`: reinterpret bits |
//! | `i32` → `i64`: zero-extend | `i64` → `i32`: wrap (mod 2³²) |
//! | `f32` → `i64`: reinterpret to i32 | `i64` → `f32`: wrap then reinterpret |
//! | `f64` → `i64`: reinterpret bits | `i64` → `f64`: reinterpret bits |
//!
//! Note the asymmetry that is easy to get wrong: `f32` → `i64` reinterprets to
//! a **32-bit** pattern and leaves the top half zero; it does NOT widen the
//! float to `f64` first. Lifting mirrors that by wrapping to i32 BEFORE
//! reinterpreting.

use crate::canon_flat::{CoreType, flatten_type, flatten_types};
use crate::canon_value::{CanonError, Realloc};
use crate::component::ValType;
use crate::value::Value;

/// One core value in transit — the flat representation's alphabet.
///
/// Kept as a small enum rather than reusing [`Value`] because the spec's
/// coercions are defined on BIT PATTERNS of a known width, and `Value`'s
/// numeric tower would silently convert where the spec reinterprets.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum CoreValue {
    I32(u32),
    I64(u64),
    F32(f32),
    F64(f64),
}

impl CoreValue {
    fn core_type(self) -> CoreType {
        match self {
            CoreValue::I32(_) => CoreType::I32,
            CoreValue::I64(_) => CoreType::I64,
            CoreValue::F32(_) => CoreType::F32,
            CoreValue::F64(_) => CoreType::F64,
        }
    }
}

/// `CoreValueIter` — walks a flat list, one core value at a time.
///
/// `done()` matters: the spec asserts every value was consumed, and a lift
/// that stops early has silently dropped a parameter.
pub struct CoreValueIter<'a> {
    values: &'a [CoreValue],
    i: usize,
}

impl<'a> CoreValueIter<'a> {
    pub fn new(values: &'a [CoreValue]) -> Self {
        CoreValueIter { values, i: 0 }
    }

    fn next(&mut self) -> Result<CoreValue, CanonError> {
        let v = self
            .values
            .get(self.i)
            .copied()
            .ok_or(CanonError::Unsupported("flat lift ran off the end of its values"))?;
        self.i += 1;
        Ok(v)
    }

    pub fn done(&self) -> bool {
        self.i == self.values.len()
    }
}

// ── coercions ─────────────────────────────────────────────────────────────
//
// The spec's `encode_float_as_i32` / `decode_i32_as_float` and friends. These
// are BIT REINTERPRETATIONS, never numeric conversions: `1.0f32` becomes
// `0x3F800000`, not `1`.

fn encode_f32_as_i32(f: f32) -> u32 {
    f.to_bits()
}

fn decode_i32_as_f32(i: u32) -> f32 {
    f32::from_bits(i)
}

fn encode_f64_as_i64(f: f64) -> u64 {
    f.to_bits()
}

fn decode_i64_as_f64(i: u64) -> f64 {
    f64::from_bits(i)
}

/// `wrap_i64_to_i32(i)` — `i % 2**32`.
fn wrap_i64_to_i32(i: u64) -> u32 {
    i as u32
}

/// Coerce on the LOWER side: a value that flattened as `have` must travel in a
/// `want` slot, because `join` widened the position for some other case.
fn coerce_lower(v: CoreValue, have: CoreType, want: CoreType) -> Result<CoreValue, CanonError> {
    if have == want {
        return Ok(v);
    }
    Ok(match (have, want, v) {
        (CoreType::F32, CoreType::I32, CoreValue::F32(f)) => CoreValue::I32(encode_f32_as_i32(f)),
        (CoreType::I32, CoreType::I64, CoreValue::I32(i)) => CoreValue::I64(i as u64),
        // 32-bit pattern in the low half — NOT a widening to f64 first.
        (CoreType::F32, CoreType::I64, CoreValue::F32(f)) => {
            CoreValue::I64(encode_f32_as_i32(f) as u64)
        }
        (CoreType::F64, CoreType::I64, CoreValue::F64(f)) => CoreValue::I64(encode_f64_as_i64(f)),
        _ => return Err(CanonError::Unsupported("flat lower: unjoinable core types")),
    })
}

/// Coerce on the LIFT side — the exact inverse of [`coerce_lower`].
fn coerce_lift(v: CoreValue, have: CoreType, want: CoreType) -> Result<CoreValue, CanonError> {
    if have == want {
        return Ok(v);
    }
    Ok(match (have, want, v) {
        (CoreType::I32, CoreType::F32, CoreValue::I32(i)) => CoreValue::F32(decode_i32_as_f32(i)),
        (CoreType::I64, CoreType::I32, CoreValue::I64(i)) => CoreValue::I32(wrap_i64_to_i32(i)),
        // Wrap FIRST, then reinterpret — mirroring the lower side.
        (CoreType::I64, CoreType::F32, CoreValue::I64(i)) => {
            CoreValue::F32(decode_i32_as_f32(wrap_i64_to_i32(i)))
        }
        (CoreType::I64, CoreType::F64, CoreValue::I64(i)) => CoreValue::F64(decode_i64_as_f64(i)),
        _ => return Err(CanonError::Unsupported("flat lift: unjoinable core types")),
    })
}

// ── lifting ───────────────────────────────────────────────────────────────

/// `lift_flat(cx, vi, t)` — one component value out of the flat stream.
pub fn lift_flat(
    memory: &crate::shared_memory::SharedMemory,
    vi: &mut CoreValueIter<'_>,
    t: &ValType,
    ptr_type: CoreType,
) -> Result<Value, CanonError> {
    Ok(match t {
        ValType::Bool => Value::Bool(matches!(vi.next()?, CoreValue::I32(i) if i != 0)),
        // ⛔ The narrow widths arrive in a full core `i32` slot and must be
        // NARROWED on the way in — `CanonicalABI.md:3282` lifts each through
        // its own width, so the high bits of the slot are not part of the
        // value. Signed ones then sign-extend, unsigned ones do not: the same
        // byte `0xFF` is `-1` as `s8` and `255` as `u8`. Taking the slot
        // verbatim would carry a silently out-of-range value.
        ValType::S8 => match vi.next()? {
            CoreValue::I32(i) => Value::I32(i as u8 as i8 as i32),
            other => return Err(mismatch(other, CoreType::I32)),
        },
        ValType::U8 => match vi.next()? {
            CoreValue::I32(i) => Value::I32(i as u8 as i32),
            other => return Err(mismatch(other, CoreType::I32)),
        },
        ValType::S16 => match vi.next()? {
            CoreValue::I32(i) => Value::I32(i as u16 as i16 as i32),
            other => return Err(mismatch(other, CoreType::I32)),
        },
        ValType::U16 => match vi.next()? {
            CoreValue::I32(i) => Value::I32(i as u16 as i32),
            other => return Err(mismatch(other, CoreType::I32)),
        },
        ValType::I32 => match vi.next()? {
            CoreValue::I32(i) => Value::I32(i as i32),
            other => return Err(mismatch(other, CoreType::I32)),
        },
        ValType::I64 => match vi.next()? {
            CoreValue::I64(i) => Value::I64(i as i64),
            other => return Err(mismatch(other, CoreType::I64)),
        },
        ValType::F32 => match vi.next()? {
            CoreValue::F32(f) => Value::F64(f as f64),
            other => return Err(mismatch(other, CoreType::F32)),
        },
        ValType::F64 => match vi.next()? {
            CoreValue::F64(f) => Value::F64(f),
            other => return Err(mismatch(other, CoreType::F64)),
        },
        // `convert_i32_to_char` — the trap is what makes this a `char` and not
        // a `u32`, so it runs here as well as on the memory path.
        ValType::Char => match vi.next()? {
            CoreValue::I32(i) => {
                let c = crate::canon_value::scalar_to_char(i)?;
                Value::String(std::sync::Arc::from(c.to_string().as_str()))
            }
            other => return Err(mismatch(other, CoreType::I32)),
        },
        // `lift_flat_flags` — one `i32` however wide the packed integer is in
        // memory, unpacked into a record of label → bool.
        ValType::Flags(labels) => match vi.next()? {
            CoreValue::I32(i) => crate::canon_value::unpack_flags(i, labels),
            other => return Err(mismatch(other, CoreType::I32)),
        },
        // 🔧 `lift_flat_list` with a length present: N elements read straight
        // from the flat sequence, no (ptr, len) pair and no memory access.
        ValType::ListFixed(elem, n) => {
            let mut items = Vec::with_capacity(*n as usize);
            for _ in 0..*n {
                items.push(lift_flat(memory, vi, elem, ptr_type)?);
            }
            Value::Object(crate::heap::alloc(crate::value::Object::new_array(items)))
        }
        // `lift_flat_string` — the (ptr, length) pair, bytes still in memory.
        ValType::String | ValType::List(_) => {
            let at = flat_ptr(vi.next()?)?;
            let len = flat_ptr(vi.next()?)?;
            // Both forms are a (ptr, length) pair, and `canon_value::load`
            // already knows how to read each from memory — so the pair is
            // written into a scratch pair-shaped read rather than duplicating
            // the decode here.
            return load_pair(memory, t, at, len);
        }
        ValType::Record(fields) => {
            let mut object = crate::value::Object::new();
            for (name, field_ty) in fields {
                let v = lift_flat(memory, vi, field_ty, ptr_type)?;
                object.properties.insert(name.as_str().into(), v);
            }
            Value::Object(crate::heap::alloc(object))
        }
        ValType::Option(inner) => {
            let cases = [(String::from("none"), None), (String::from("some"), Some((**inner).clone()))];
            let (case, payload) = lift_flat_variant(memory, vi, &cases, ptr_type)?;
            if case == 0 { Value::Null } else { payload }
        }
        ValType::Result(ok, err) => {
            let cases = [
                (String::from("ok"), ok.as_deref().cloned()),
                (String::from("error"), err.as_deref().cloned()),
            ];
            let (case, payload) = lift_flat_variant(memory, vi, &cases, ptr_type)?;
            if case == 0 {
                payload
            } else {
                let mut object = crate::value::Object::new();
                object.properties.insert("__wasi_error".into(), payload);
                Value::Object(crate::heap::alloc(object))
            }
        }
        ValType::Variant(cases) => {
            let (case, payload) = lift_flat_variant(memory, vi, cases, ptr_type)?;
            match cases.get(case as usize) {
                // Same rule `canon_value::load` follows: a case that carried no
                // payload lifts as the bare NAME, so the two directions agree.
                Some((name, _)) if matches!(payload, Value::Null) => {
                    Value::String(std::sync::Arc::from(name.as_str()))
                }
                Some((name, _)) => {
                    let mut object = crate::value::Object::new();
                    object
                        .properties
                        .insert("tag".into(), Value::String(std::sync::Arc::from(name.as_str())));
                    object.properties.insert("val".into(), payload);
                    Value::Object(crate::heap::alloc(object))
                }
                None => {
                    return Err(CanonError::DiscriminantOutOfRange {
                        got: case,
                        cases: cases.len(),
                    })
                }
            }
        }
        // A handle or async end is its index.
        ValType::Own(_)
        | ValType::Borrow(_)
        | ValType::Stream(_)
        | ValType::Future(_)
        | ValType::ErrorContext => {
            Value::I32(flat_ptr(vi.next()?)? as i32)
        }
        ValType::Any => return Err(CanonError::Unsupported("any (not a component type)")),
    })
}

/// `lift_flat_variant` — read the discriminant, lift the live case's payload
/// through the coercions, then CONSUME the remaining joined slots.
///
/// That last step is not bookkeeping: the flattening is the same width for
/// every case, so a shorter case leaves slots the caller still wrote. Skipping
/// them leaves the iterator pointing at padding, and the next parameter lifts
/// from the wrong place.
fn lift_flat_variant(
    memory: &crate::shared_memory::SharedMemory,
    vi: &mut CoreValueIter<'_>,
    cases: &[(String, Option<ValType>)],
    ptr_type: CoreType,
) -> Result<(u32, Value), CanonError> {
    let payloads: Vec<Option<ValType>> = cases.iter().map(|(_, t)| t.clone()).collect();
    let joined = crate::canon_flat::flatten_type(
        &ValType::Variant(cases.to_vec()),
        ptr_type,
    );
    // Element 0 is the discriminant; the rest are the joined payload slots.
    let joined_payload = &joined[1..];

    let case_index = flat_ptr(vi.next()?)?;
    if case_index as usize >= cases.len() {
        return Err(CanonError::DiscriminantOutOfRange {
            got: case_index,
            cases: cases.len(),
        });
    }

    let mut consumed = 0usize;
    let value = match &payloads[case_index as usize] {
        None => Value::Null,
        Some(t) => {
            let want = flatten_type(t, ptr_type);
            let mut coerced: Vec<CoreValue> = Vec::with_capacity(want.len());
            for w in &want {
                let have = joined_payload
                    .get(consumed)
                    .copied()
                    .ok_or(CanonError::Unsupported("variant flattening too short"))?;
                let raw = vi.next()?;
                coerced.push(coerce_lift(raw, have, *w)?);
                consumed += 1;
            }
            let mut inner = CoreValueIter::new(&coerced);
            lift_flat(memory, &mut inner, t, ptr_type)?
        }
    };

    // Drain the slots this case did not use.
    for _ in consumed..joined_payload.len() {
        let _ = vi.next()?;
    }
    Ok((case_index, value))
}

// ── lowering ──────────────────────────────────────────────────────────────

/// `lower_flat(cx, v, t)` — one component value into core values.
pub fn lower_flat(
    memory: &crate::shared_memory::SharedMemory,
    realloc: &mut Realloc<'_>,
    v: &Value,
    t: &ValType,
    ptr_type: CoreType,
) -> Result<Vec<CoreValue>, CanonError> {
    Ok(match t {
        ValType::Bool => vec![CoreValue::I32(u32::from(v.as_bool()))],
        // Lowering narrows to the declared width and then widens back into the
        // i32 slot, so a value out of range for its type cannot be smuggled
        // through in the slot's spare bits.
        ValType::S8 => vec![CoreValue::I32(v.as_i32() as i8 as i32 as u32)],
        ValType::U8 => vec![CoreValue::I32(v.as_i32() as u8 as u32)],
        ValType::S16 => vec![CoreValue::I32(v.as_i32() as i16 as i32 as u32)],
        ValType::U16 => vec![CoreValue::I32(v.as_i32() as u16 as u32)],
        ValType::I32 => vec![CoreValue::I32(v.as_i32() as u32)],
        ValType::I64 => vec![CoreValue::I64(v.as_i64() as u64)],
        ValType::F32 => vec![CoreValue::F32(v.as_f64() as f32)],
        ValType::F64 => vec![CoreValue::F64(v.as_f64())],
        ValType::Char => vec![CoreValue::I32(crate::canon_value::value_to_scalar(v)?)],
        // `lower_flat_flags` — `[pack_flags_into_int(v, labels)]`.
        ValType::Flags(labels) => vec![CoreValue::I32(crate::canon_value::pack_flags(v, labels))],
        // 🔧 N elements' flat values in sequence. The count check is the same
        // assertion `store` makes: the length is part of the TYPE, so a
        // mismatched value cannot be silently padded or truncated.
        ValType::ListFixed(elem, n) => {
            let items = crate::canon_value::array_items_public(v);
            if items.len() as u32 != *n {
                return Err(CanonError::FixedListLength {
                    got: items.len(),
                    want: *n,
                });
            }
            let mut flat = Vec::new();
            for item in &items {
                flat.extend(lower_flat(memory, realloc, item, elem, ptr_type)?);
            }
            flat
        }
        // The bytes go to memory; only the pair travels flat.
        ValType::String | ValType::List(_) => {
            let (at, len) = store_pair(memory, realloc, v, t)?;
            vec![flat_from_ptr(at, ptr_type), flat_from_ptr(len, ptr_type)]
        }
        ValType::Record(fields) => {
            let mut flat = Vec::new();
            for (name, field_ty) in fields {
                let fv = crate::canon_value::record_field_public(v, name);
                flat.extend(lower_flat(memory, realloc, &fv, field_ty, ptr_type)?);
            }
            flat
        }
        ValType::Option(inner) => {
            let cases = [(String::from("none"), None), (String::from("some"), Some((**inner).clone()))];
            let (idx, payload) = if matches!(v, Value::Null) {
                (0u32, Value::Null)
            } else {
                (1u32, v.clone())
            };
            lower_flat_variant(memory, realloc, idx, &payload, &cases, ptr_type)?
        }
        ValType::Result(ok, err) => {
            let cases = [
                (String::from("ok"), ok.as_deref().cloned()),
                (String::from("error"), err.as_deref().cloned()),
            ];
            let (idx, payload) = crate::canon_value::result_parts_public(v);
            lower_flat_variant(memory, realloc, idx, &payload, &cases, ptr_type)?
        }
        ValType::Variant(cases) => {
            let (idx, payload) = crate::canon_value::variant_case_public(v, cases)?;
            lower_flat_variant(memory, realloc, idx, &payload, cases, ptr_type)?
        }
        ValType::Own(_)
        | ValType::Borrow(_)
        | ValType::Stream(_)
        | ValType::Future(_)
        | ValType::ErrorContext => {
            vec![CoreValue::I32(v.as_i32() as u32)]
        }
        ValType::Any => return Err(CanonError::Unsupported("any (not a component type)")),
    })
}

/// `lower_flat_variant` — discriminant, coerced payload, then PAD to the joined
/// width with zeroes of the right core type.
///
/// The padding is required, not defensive: every case must occupy the same
/// number of core values, or the receiver's `lift_flat_variant` reads the next
/// parameter as this variant's tail.
fn lower_flat_variant(
    memory: &crate::shared_memory::SharedMemory,
    realloc: &mut Realloc<'_>,
    case_index: u32,
    payload: &Value,
    cases: &[(String, Option<ValType>)],
    ptr_type: CoreType,
) -> Result<Vec<CoreValue>, CanonError> {
    let joined = crate::canon_flat::flatten_type(&ValType::Variant(cases.to_vec()), ptr_type);
    let joined_payload = &joined[1..];

    let mut flat = vec![CoreValue::I32(case_index)];
    let mut written = 0usize;

    if let Some((_, Some(t))) = cases.get(case_index as usize) {
        let have_types = flatten_type(t, ptr_type);
        let lowered = lower_flat(memory, realloc, payload, t, ptr_type)?;
        for (i, fv) in lowered.into_iter().enumerate() {
            let have = have_types
                .get(i)
                .copied()
                .unwrap_or_else(|| fv.core_type());
            let want = joined_payload
                .get(written)
                .copied()
                .ok_or(CanonError::Unsupported("variant flattening too short"))?;
            flat.push(coerce_lower(fv, have, want)?);
            written += 1;
        }
    }

    for want in &joined_payload[written..] {
        flat.push(match want {
            CoreType::I32 => CoreValue::I32(0),
            CoreType::I64 => CoreValue::I64(0),
            CoreType::F32 => CoreValue::F32(0.0),
            CoreType::F64 => CoreValue::F64(0.0),
        });
    }
    Ok(flat)
}

/// `lower_flat_values` — a whole parameter list, spilling to memory when it
/// exceeds `max_flat`.
///
/// The spill is the reason this is not just `map(lower_flat)`: past the limit
/// the ABI stores a TUPLE in linear memory and passes one pointer, and the
/// receiver must agree about which happened. `canon_flat::flatten_functype`
/// decides that from the same counts.
pub fn lower_flat_values(
    memory: &crate::shared_memory::SharedMemory,
    realloc: &mut Realloc<'_>,
    max_flat: usize,
    values: &[Value],
    types: &[ValType],
    ptr_type: CoreType,
) -> Result<Vec<CoreValue>, CanonError> {
    if values.len() != types.len() {
        return Err(CanonError::Unsupported(
            "lower_flat_values: value/type count mismatch",
        ));
    }
    let flat_types = flatten_types(types, ptr_type);
    if flat_types.len() <= max_flat {
        let mut flat = Vec::new();
        for (v, t) in values.iter().zip(types) {
            flat.extend(lower_flat(memory, realloc, v, t, ptr_type)?);
        }
        return Ok(flat);
    }

    // Over the limit: store the arguments as a tuple in memory and pass the
    // pointer. The tuple is a record of the parameter types, so the layout is
    // the one `canon_value` already computes.
    let tuple_ty = ValType::Record(
        types
            .iter()
            .enumerate()
            .map(|(i, t)| (format!("{i}"), t.clone()))
            .collect(),
    );
    let mut object = crate::value::Object::new();
    for (i, v) in values.iter().enumerate() {
        object.properties.insert(format!("{i}").into(), v.clone());
    }
    let tuple = Value::Object(crate::heap::alloc(object));

    let size = crate::canon_layout::elem_size(&tuple_ty);
    let align = crate::canon_layout::alignment(&tuple_ty);
    let at = realloc(size, align).ok_or(CanonError::OutOfMemory { size })?;
    crate::canon_value::store_with(memory, realloc, &tuple, &tuple_ty, at)?;
    Ok(vec![flat_from_ptr(at, ptr_type)])
}

/// `lift_flat_values` — the mirror, unpacking a spilled tuple when the flat
/// form is a single pointer.
pub fn lift_flat_values(
    memory: &crate::shared_memory::SharedMemory,
    max_flat: usize,
    values: &[CoreValue],
    types: &[ValType],
    ptr_type: CoreType,
) -> Result<Vec<Value>, CanonError> {
    let flat_types = flatten_types(types, ptr_type);
    if flat_types.len() <= max_flat {
        let mut vi = CoreValueIter::new(values);
        let mut out = Vec::with_capacity(types.len());
        for t in types {
            out.push(lift_flat(memory, &mut vi, t, ptr_type)?);
        }
        return Ok(out);
    }

    let at = values
        .first()
        .copied()
        .ok_or(CanonError::Unsupported("lift_flat_values: no pointer"))?;
    let at = flat_ptr(at)?;
    let tuple_ty = ValType::Record(
        types
            .iter()
            .enumerate()
            .map(|(i, t)| (format!("{i}"), t.clone()))
            .collect(),
    );
    let tuple = crate::canon_value::load(memory, &tuple_ty, at)?;
    let mut out = Vec::with_capacity(types.len());
    for i in 0..types.len() {
        out.push(crate::canon_value::record_field_public(&tuple, &format!("{i}")));
    }
    Ok(out)
}

// ── helpers ───────────────────────────────────────────────────────────────

fn mismatch(v: CoreValue, want: CoreType) -> CanonError {
    CanonError::FlatTypeMismatch {
        got: v.core_type(),
        want,
    }
}

/// A pointer-or-length core value as a `u32`, whatever width it travelled in.
fn flat_ptr(v: CoreValue) -> Result<u32, CanonError> {
    Ok(match v {
        CoreValue::I32(i) => i,
        CoreValue::I64(i) => i as u32,
        _ => return Err(CanonError::Unsupported("flat: expected an integer slot")),
    })
}

fn flat_from_ptr(p: u32, ptr_type: CoreType) -> CoreValue {
    match ptr_type {
        CoreType::I64 => CoreValue::I64(p as u64),
        _ => CoreValue::I32(p),
    }
}

/// Read a (ptr, length)-shaped value by writing the pair into a scratch cell
/// and letting `canon_value::load` decode it — the spec's "reuse the previous
/// definitions" for strings and lists.
fn load_pair(
    memory: &crate::shared_memory::SharedMemory,
    t: &ValType,
    at: u32,
    len: u32,
) -> Result<Value, CanonError> {
    crate::canon_value::load_pair_public(memory, t, at, len)
}

fn store_pair(
    memory: &crate::shared_memory::SharedMemory,
    realloc: &mut Realloc<'_>,
    v: &Value,
    t: &ValType,
) -> Result<(u32, u32), CanonError> {
    crate::canon_value::store_pair_public(memory, realloc, v, t)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::shared_memory::SharedMemory;

    fn s(n: &str) -> String {
        n.to_string()
    }

    fn mem() -> SharedMemory {
        SharedMemory::new(4096)
    }

    /// A bump allocator over the tail of the test's memory.
    fn with_alloc<R>(m: &SharedMemory, f: impl FnOnce(&mut Realloc<'_>) -> R) -> R {
        let mut next = 512u32;
        let mut alloc = |size: u32, align: u32| -> Option<u32> {
            let at = crate::canon_layout::align_to(next, align.max(1));
            next = at + size;
            (next as usize <= m.len()).then_some(at)
        };
        let mut r: Realloc<'_> = &mut alloc;
        f(&mut r)
    }

    fn round_trip(v: &Value, t: &ValType) -> Value {
        let m = mem();
        let flat = with_alloc(&m, |r| lower_flat(&m, r, v, t, CoreType::I32).unwrap());
        let mut vi = CoreValueIter::new(&flat);
        let back = lift_flat(&m, &mut vi, t, CoreType::I32).unwrap();
        assert!(vi.done(), "lift must consume every core value it was given");
        back
    }

    #[test]
    fn scalars_round_trip_flat() {
        assert_eq!(round_trip(&Value::I32(-7), &ValType::I32).as_i32(), -7);
        assert_eq!(round_trip(&Value::I64(-9), &ValType::I64).as_i64(), -9);
        assert_eq!(round_trip(&Value::F64(1.5), &ValType::F64).as_f64(), 1.5);
        assert!(round_trip(&Value::Bool(true), &ValType::Bool).as_bool());
    }

    /// A string's BYTES stay in memory; only (ptr, length) travels flat. The
    /// round trip therefore proves the pair was written and read at the same
    /// widths — the commonest way to halve a signature.
    #[test]
    fn a_string_travels_as_a_pair_and_survives() {
        let v = Value::String("héllo".into());
        let back = round_trip(&v, &ValType::String);
        assert_eq!(format!("{back}"), "héllo");
    }

    #[test]
    fn a_record_round_trips_field_by_field() {
        let t = ValType::Record(vec![
            (s("a"), ValType::I32),
            (s("b"), ValType::String),
            (s("c"), ValType::F64),
        ]);
        let mut o = crate::value::Object::new();
        o.properties.insert("a".into(), Value::I32(3));
        o.properties.insert("b".into(), Value::String("x".into()));
        o.properties.insert("c".into(), Value::F64(2.5));
        let v = Value::Object(crate::heap::alloc(o));

        let back = round_trip(&v, &t);
        assert_eq!(crate::canon_value::record_field_public(&back, "a").as_i32(), 3);
        assert_eq!(
            format!("{}", crate::canon_value::record_field_public(&back, "b")),
            "x"
        );
        assert_eq!(
            crate::canon_value::record_field_public(&back, "c").as_f64(),
            2.5
        );
    }

    /// ▶▶THE CASE THE COERCION TABLES EXIST FOR.
    ///
    /// `join(i32, f64) = i64`, so BOTH cases travel in an `i64` slot: the
    /// `i32` case zero-extended, the `f64` case bit-reinterpreted. Getting the
    /// direction wrong turns `2.5` into `4612811918334230528` — a plausible
    /// number, silently.
    #[test]
    fn a_variant_survives_the_join_widening_both_ways() {
        let t = ValType::Variant(vec![
            (s("small"), Some(ValType::I32)),
            (s("wide"), Some(ValType::F64)),
        ]);

        let mut small = crate::value::Object::new();
        small.properties.insert("tag".into(), Value::String("small".into()));
        small.properties.insert("val".into(), Value::I32(42));
        let small = Value::Object(crate::heap::alloc(small));
        let back = round_trip(&small, &t);
        assert_eq!(
            format!("{}", crate::canon_value::record_field_public(&back, "tag")),
            "small"
        );
        assert_eq!(
            crate::canon_value::record_field_public(&back, "val").as_i32(),
            42,
            "an i32 case zero-extended into an i64 slot must come back as 42"
        );

        let mut wide = crate::value::Object::new();
        wide.properties.insert("tag".into(), Value::String("wide".into()));
        wide.properties.insert("val".into(), Value::F64(2.5));
        let wide = Value::Object(crate::heap::alloc(wide));
        let back = round_trip(&wide, &t);
        assert_eq!(
            crate::canon_value::record_field_public(&back, "val").as_f64(),
            2.5,
            "an f64 case REINTERPRETED into an i64 slot must come back as 2.5, not its bit pattern"
        );
    }

    /// Every case must occupy the same number of core values. A short case is
    /// PADDED on lower and DRAINED on lift; if either side forgets, the next
    /// parameter is read from this variant's tail.
    #[test]
    fn a_payload_free_case_still_occupies_the_joined_width() {
        let t = ValType::Variant(vec![
            (s("none"), None),
            (s("wide"), Some(ValType::F64)),
        ]);
        let m = mem();
        let bare = Value::String("none".into());
        let flat = with_alloc(&m, |r| lower_flat(&m, r, &bare, &t, CoreType::I32).unwrap());
        // discriminant + one joined payload slot, even though this case has no
        // payload at all.
        assert_eq!(flat.len(), 2, "a payload-free case must still be padded");

        let mut vi = CoreValueIter::new(&flat);
        let back = lift_flat(&m, &mut vi, &t, CoreType::I32).unwrap();
        assert!(vi.done(), "the unused slot must be drained, not left behind");
        assert_eq!(format!("{back}"), "none");
    }

    /// `result<_, error-code>` — the shape every WASI 0.3.1 call returns.
    #[test]
    fn a_result_error_arm_round_trips_flat() {
        let t = ValType::Result(None, Some(Box::new(ValType::String)));
        let mut e = crate::value::Object::new();
        e.properties
            .insert("__wasi_error".into(), Value::String("no-entry".into()));
        let v = Value::Object(crate::heap::alloc(e));

        let back = round_trip(&v, &t);
        assert_eq!(
            format!(
                "{}",
                crate::canon_value::record_field_public(&back, "__wasi_error")
            ),
            "no-entry",
            "an error arm must not round-trip into an ok"
        );
    }

    /// Past `MAX_FLAT_PARAMS` the whole argument list becomes ONE pointer to a
    /// tuple in memory, and the receiver has to know that happened. Both sides
    /// decide from the same flattening, so this proves they agree.
    #[test]
    fn an_oversized_argument_list_spills_to_memory_and_comes_back() {
        let types: Vec<ValType> = (0..crate::canon_flat::MAX_FLAT_PARAMS + 2)
            .map(|_| ValType::I32)
            .collect();
        let values: Vec<Value> = (0..types.len()).map(|i| Value::I32(i as i32)).collect();

        let m = mem();
        let flat = with_alloc(&m, |r| {
            lower_flat_values(
                &m,
                r,
                crate::canon_flat::MAX_FLAT_PARAMS,
                &values,
                &types,
                CoreType::I32,
            )
            .unwrap()
        });
        assert_eq!(flat.len(), 1, "over the limit, exactly one pointer travels");

        let back = lift_flat_values(
            &m,
            crate::canon_flat::MAX_FLAT_PARAMS,
            &flat,
            &types,
            CoreType::I32,
        )
        .unwrap();
        assert_eq!(back.len(), types.len());
        for (i, v) in back.iter().enumerate() {
            assert_eq!(v.as_i32(), i as i32, "argument {i} survived the spill");
        }
    }

    /// Under the limit nothing spills — the values travel as themselves.
    #[test]
    fn a_small_argument_list_stays_flat() {
        let types = [ValType::I32, ValType::I64];
        let values = [Value::I32(1), Value::I64(2)];
        let m = mem();
        let flat = with_alloc(&m, |r| {
            lower_flat_values(
                &m,
                r,
                crate::canon_flat::MAX_FLAT_PARAMS,
                &values,
                &types,
                CoreType::I32,
            )
            .unwrap()
        });
        assert_eq!(flat, [CoreValue::I32(1), CoreValue::I64(2)]);
    }
}
