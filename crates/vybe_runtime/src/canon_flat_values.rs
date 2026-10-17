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

use crate::canon_flat::{CoreType, flat_count_bounded, flat_types_count_bounded, flatten_type};
use crate::canon_value::{CanonError, Realloc};
use crate::component::ValType;
use crate::value::Value;
use std::collections::HashMap;

#[derive(Debug)]
struct VariantLayout {
    joined: Vec<CoreType>,
    cases: Vec<Vec<CoreType>>,
}

/// Layouts belong to one immutable function type. Keys are addresses of its
/// nested `ValType`s; the owning boxed function type must outlive this cache.
#[derive(Debug, Default)]
pub(crate) struct FlatLayoutCache {
    variants: HashMap<usize, VariantLayout>,
    param_tuple: Option<(ValType, crate::canon_layout::MemLayout)>,
    result_tuple: Option<(ValType, crate::canon_layout::MemLayout)>,
    pub params_flat_len: usize,
    pub result_flat_len: usize,
}

impl FlatLayoutCache {
    pub(crate) fn for_functype(ft: &crate::canon_def::CanonFuncType, ptr: CoreType) -> Self {
        let mut cache = Self {
            variants: HashMap::new(),
            param_tuple: None,
            result_tuple: None,
            params_flat_len: flat_types_count_bounded(
                &ft.params,
                crate::canon_flat::MAX_FLAT_PARAMS,
            ),
            result_flat_len: ft.result.as_ref().map_or(0, |t| {
                flat_count_bounded(t, crate::canon_flat::MAX_FLAT_PARAMS)
            }),
        };
        if cache.params_flat_len > crate::canon_flat::MAX_FLAT_PARAMS {
            cache.param_tuple = Some(Self::tuple_layout(&ft.params));
        }
        if cache.result_flat_len > crate::canon_flat::MAX_FLAT_RESULTS {
            cache.result_tuple = Some(Self::tuple_layout(
                &ft.result.iter().cloned().collect::<Vec<_>>(),
            ));
        }
        for ty in ft.params.iter().chain(ft.result.iter()) {
            cache.collect(ty, ptr);
        }
        cache
    }

    fn tuple_layout(types: &[ValType]) -> (ValType, crate::canon_layout::MemLayout) {
        let ty = ValType::Record(
            types
                .iter()
                .enumerate()
                .map(|(i, t)| (format!("{i}"), t.clone()))
                .collect(),
        );
        let layout = crate::canon_layout::MemLayout::new(&ty);
        (ty, layout)
    }

    fn tuple_for(&self, max_flat: usize) -> Option<&(ValType, crate::canon_layout::MemLayout)> {
        if max_flat == crate::canon_flat::MAX_FLAT_PARAMS {
            self.param_tuple.as_ref()
        } else {
            self.result_tuple.as_ref()
        }
    }

    fn collect(&mut self, ty: &ValType, ptr: CoreType) {
        match ty {
            ValType::Option(inner) => {
                self.collect(inner, ptr);
                self.insert_variant(ty, [None, Some(inner.as_ref())].into_iter(), ptr);
            }
            ValType::Result(ok, err) => {
                for inner in [ok.as_deref(), err.as_deref()].into_iter().flatten() {
                    self.collect(inner, ptr);
                }
                self.insert_variant(ty, [ok.as_deref(), err.as_deref()].into_iter(), ptr);
            }
            ValType::Variant(cases) => {
                for (_, inner) in cases {
                    if let Some(inner) = inner {
                        self.collect(inner, ptr);
                    }
                }
                self.insert_variant(ty, cases.iter().map(|(_, inner)| inner.as_ref()), ptr);
            }
            ValType::Record(fields) => {
                for (_, inner) in fields {
                    self.collect(inner, ptr);
                }
            }
            ValType::ListFixed(inner, _) => self.collect(inner, ptr),
            _ => {}
        }
    }

    fn insert_variant<'a>(
        &mut self,
        ty: &ValType,
        cases: impl ExactSizeIterator<Item = Option<&'a ValType>> + Clone,
        ptr: CoreType,
    ) {
        let joined = crate::canon_flat::flatten_variant_cases(cases.clone(), ptr);
        let cases = cases
            .map(|case| case.map_or_else(Vec::new, |inner| flatten_type(inner, ptr)))
            .collect();
        self.variants.insert(
            ty as *const ValType as usize,
            VariantLayout { joined, cases },
        );
    }

    fn variant(&self, ty: &ValType) -> Option<&VariantLayout> {
        self.variants.get(&(ty as *const ValType as usize))
    }
}

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
            .ok_or(CanonError::Unsupported(
                "flat lift ran off the end of its values",
            ))?;
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
    lift_flat_impl(memory, vi, t, ptr_type, None)
}

fn lift_flat_impl(
    memory: &crate::shared_memory::SharedMemory,
    vi: &mut CoreValueIter<'_>,
    t: &ValType,
    ptr_type: CoreType,
    layouts: Option<&FlatLayoutCache>,
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
        ValType::U32 => match vi.next()? {
            CoreValue::I32(i) => Value::F64(i as f64),
            other => return Err(mismatch(other, CoreType::I32)),
        },
        ValType::I64 => match vi.next()? {
            CoreValue::I64(i) => Value::I64(i as i64),
            other => return Err(mismatch(other, CoreType::I64)),
        },
        ValType::U64 => match vi.next()? {
            CoreValue::I64(i) => Value::bigint_u64(i),
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
                items.push(lift_flat_impl(memory, vi, elem, ptr_type, layouts)?);
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
                let v = lift_flat_impl(memory, vi, field_ty, ptr_type, layouts)?;
                object.properties.insert(name.as_str().into(), v);
            }
            Value::Object(crate::heap::alloc(object))
        }
        ValType::Option(inner) => {
            let cases = [None, Some(inner.as_ref())];
            let (case, payload) = lift_flat_variant(
                memory,
                vi,
                cases.iter().copied(),
                ptr_type,
                layouts.and_then(|c| c.variant(t)),
                layouts,
            )?;
            if case == 0 { Value::Null } else { payload }
        }
        ValType::Result(ok, err) => {
            let cases = [ok.as_deref(), err.as_deref()];
            let (case, payload) = lift_flat_variant(
                memory,
                vi,
                cases.iter().copied(),
                ptr_type,
                layouts.and_then(|c| c.variant(t)),
                layouts,
            )?;
            if case == 0 {
                payload
            } else {
                let mut object = crate::value::Object::new();
                object.properties.insert("__wasi_error".into(), payload);
                Value::Object(crate::heap::alloc(object))
            }
        }
        ValType::Variant(cases) => {
            let (case, payload) = lift_flat_variant(
                memory,
                vi,
                cases.iter().map(|(_, payload)| payload.as_ref()),
                ptr_type,
                layouts.and_then(|c| c.variant(t)),
                layouts,
            )?;
            match cases.get(case as usize) {
                // Same rule `canon_value::load` follows: a case that carried no
                // payload lifts as the bare NAME, so the two directions agree.
                Some((name, _)) if matches!(payload, Value::Null) => {
                    Value::String(std::sync::Arc::from(name.as_str()))
                }
                Some((name, _)) => {
                    let mut object = crate::value::Object::new();
                    object.properties.insert(
                        "tag".into(),
                        Value::String(std::sync::Arc::from(name.as_str())),
                    );
                    object.properties.insert("val".into(), payload);
                    Value::Object(crate::heap::alloc(object))
                }
                None => {
                    return Err(CanonError::DiscriminantOutOfRange {
                        got: case,
                        cases: cases.len(),
                    });
                }
            }
        }
        // A handle or async end is its index.
        ValType::Own(_)
        | ValType::Borrow(_)
        | ValType::Stream(_)
        | ValType::Future(_)
        | ValType::ErrorContext => Value::I32(flat_ptr(vi.next()?)? as i32),
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
fn lift_flat_variant<'a>(
    memory: &crate::shared_memory::SharedMemory,
    vi: &mut CoreValueIter<'_>,
    cases: impl ExactSizeIterator<Item = Option<&'a ValType>> + Clone,
    ptr_type: CoreType,
    layout: Option<&VariantLayout>,
    layouts: Option<&FlatLayoutCache>,
) -> Result<(u32, Value), CanonError> {
    let joined_storage;
    let joined = if let Some(layout) = layout {
        &layout.joined
    } else {
        joined_storage = crate::canon_flat::flatten_variant_cases(cases.clone(), ptr_type);
        &joined_storage
    };
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
    let value = match cases.clone().nth(case_index as usize).flatten() {
        None => Value::Null,
        Some(t) => {
            let want_storage;
            let want = if let Some(layout) = layout {
                &layout.cases[case_index as usize]
            } else {
                want_storage = flatten_type(t, ptr_type);
                &want_storage
            };
            let mut coerced: Vec<CoreValue> = Vec::with_capacity(want.len());
            for w in want {
                let have = joined_payload
                    .get(consumed)
                    .copied()
                    .ok_or(CanonError::Unsupported("variant flattening too short"))?;
                let raw = vi.next()?;
                coerced.push(coerce_lift(raw, have, *w)?);
                consumed += 1;
            }
            let mut inner = CoreValueIter::new(&coerced);
            lift_flat_impl(memory, &mut inner, t, ptr_type, layouts)?
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
    let mut flat = Vec::new();
    lower_flat_into(memory, realloc, v, t, ptr_type, None, &mut flat)?;
    Ok(flat)
}

fn lower_flat_into(
    memory: &crate::shared_memory::SharedMemory,
    realloc: &mut Realloc<'_>,
    v: &Value,
    t: &ValType,
    ptr_type: CoreType,
    layouts: Option<&FlatLayoutCache>,
    flat: &mut Vec<CoreValue>,
) -> Result<(), CanonError> {
    match t {
        ValType::Bool => flat.push(CoreValue::I32(u32::from(v.as_bool()))),
        // Lowering narrows to the declared width and then widens back into the
        // i32 slot, so a value out of range for its type cannot be smuggled
        // through in the slot's spare bits.
        ValType::S8 => flat.push(CoreValue::I32(v.as_i32() as i8 as i32 as u32)),
        ValType::U8 => flat.push(CoreValue::I32(v.as_i32() as u8 as u32)),
        ValType::S16 => flat.push(CoreValue::I32(v.as_i32() as i16 as i32 as u32)),
        ValType::U16 => flat.push(CoreValue::I32(v.as_i32() as u16 as u32)),
        ValType::I32 => flat.push(CoreValue::I32(v.as_i32() as u32)),
        ValType::U32 => flat.push(CoreValue::I32(crate::canon_value::unsigned_u32(v)?)),
        ValType::I64 => flat.push(CoreValue::I64(v.as_i64() as u64)),
        ValType::U64 => flat.push(CoreValue::I64(crate::canon_value::unsigned_u64(v)?)),
        ValType::F32 => flat.push(CoreValue::F32(v.as_f64() as f32)),
        ValType::F64 => flat.push(CoreValue::F64(v.as_f64())),
        ValType::Char => flat.push(CoreValue::I32(crate::canon_value::value_to_scalar(v)?)),
        // `lower_flat_flags` — `[pack_flags_into_int(v, labels)]`.
        ValType::Flags(labels) => {
            flat.push(CoreValue::I32(crate::canon_value::pack_flags(v, labels)))
        }
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
            for item in &items {
                lower_flat_into(memory, realloc, item, elem, ptr_type, layouts, flat)?;
            }
        }
        // The bytes go to memory; only the pair travels flat.
        ValType::String | ValType::List(_) => {
            let (at, len) = store_pair(memory, realloc, v, t)?;
            flat.push(flat_from_ptr(at, ptr_type));
            flat.push(flat_from_ptr(len, ptr_type));
        }
        ValType::Record(fields) => {
            for (name, field_ty) in fields {
                let fv = crate::canon_value::record_field_public(v, name);
                lower_flat_into(memory, realloc, &fv, field_ty, ptr_type, layouts, flat)?;
            }
        }
        ValType::Option(inner) => {
            let cases = [None, Some(inner.as_ref())];
            let (idx, payload) = if matches!(v, Value::Null) {
                (0u32, Value::Null)
            } else {
                (1u32, v.clone())
            };
            lower_flat_variant_into(
                memory,
                realloc,
                idx,
                &payload,
                cases.iter().copied(),
                ptr_type,
                layouts.and_then(|c| c.variant(t)),
                layouts,
                flat,
            )?;
        }
        ValType::Result(ok, err) => {
            let cases = [ok.as_deref(), err.as_deref()];
            let (idx, payload) = crate::canon_value::result_parts_public(v);
            lower_flat_variant_into(
                memory,
                realloc,
                idx,
                &payload,
                cases.iter().copied(),
                ptr_type,
                layouts.and_then(|c| c.variant(t)),
                layouts,
                flat,
            )?;
        }
        ValType::Variant(cases) => {
            let (idx, payload) = crate::canon_value::variant_case_public(v, cases)?;
            lower_flat_variant_into(
                memory,
                realloc,
                idx,
                &payload,
                cases.iter().map(|(_, payload)| payload.as_ref()),
                ptr_type,
                layouts.and_then(|c| c.variant(t)),
                layouts,
                flat,
            )?;
        }
        ValType::Own(_)
        | ValType::Borrow(_)
        | ValType::Stream(_)
        | ValType::Future(_)
        | ValType::ErrorContext => {
            flat.push(CoreValue::I32(v.as_i32() as u32));
        }
        ValType::Any => return Err(CanonError::Unsupported("any (not a component type)")),
    }
    Ok(())
}

/// `lower_flat_variant` — discriminant, coerced payload, then PAD to the joined
/// width with zeroes of the right core type.
///
/// The padding is required, not defensive: every case must occupy the same
/// number of core values, or the receiver's `lift_flat_variant` reads the next
/// parameter as this variant's tail.
fn lower_flat_variant_into<'a>(
    memory: &crate::shared_memory::SharedMemory,
    realloc: &mut Realloc<'_>,
    case_index: u32,
    payload: &Value,
    cases: impl ExactSizeIterator<Item = Option<&'a ValType>> + Clone,
    ptr_type: CoreType,
    layout: Option<&VariantLayout>,
    layouts: Option<&FlatLayoutCache>,
    flat: &mut Vec<CoreValue>,
) -> Result<(), CanonError> {
    let joined_storage;
    let joined = if let Some(layout) = layout {
        &layout.joined
    } else {
        joined_storage = crate::canon_flat::flatten_variant_cases(cases.clone(), ptr_type);
        &joined_storage
    };
    let joined_payload = &joined[1..];

    flat.push(CoreValue::I32(case_index));
    let payload_start = flat.len();
    let mut written = 0usize;

    if let Some(t) = cases.clone().nth(case_index as usize).flatten() {
        let have_storage;
        let have_types = if let Some(layout) = layout {
            &layout.cases[case_index as usize]
        } else {
            have_storage = flatten_type(t, ptr_type);
            &have_storage
        };
        lower_flat_into(memory, realloc, payload, t, ptr_type, layouts, flat)?;
        written = flat.len() - payload_start;
        for i in 0..written {
            let fv = flat[payload_start + i];
            let have = have_types.get(i).copied().unwrap_or_else(|| fv.core_type());
            let want = joined_payload
                .get(i)
                .copied()
                .ok_or(CanonError::Unsupported("variant flattening too short"))?;
            flat[payload_start + i] = coerce_lower(fv, have, want)?;
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
    Ok(())
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
    lower_flat_values_impl(
        memory, realloc, max_flat, values, types, ptr_type, None, None, None,
    )
}

pub(crate) fn lower_flat_values_cached(
    memory: &crate::shared_memory::SharedMemory,
    realloc: &mut Realloc<'_>,
    max_flat: usize,
    values: &[Value],
    types: &[ValType],
    ptr_type: CoreType,
    layouts: &FlatLayoutCache,
    flat_count: usize,
) -> Result<Vec<CoreValue>, CanonError> {
    lower_flat_values_impl(
        memory,
        realloc,
        max_flat,
        values,
        types,
        ptr_type,
        Some(layouts),
        Some(flat_count),
        None,
    )
}

pub(crate) fn lower_flat_values_cached_into(
    memory: &crate::shared_memory::SharedMemory,
    realloc: &mut Realloc<'_>,
    max_flat: usize,
    values: &[Value],
    types: &[ValType],
    ptr_type: CoreType,
    layouts: &FlatLayoutCache,
    flat_count: usize,
    at: u32,
) -> Result<Vec<CoreValue>, CanonError> {
    lower_flat_values_impl(
        memory,
        realloc,
        max_flat,
        values,
        types,
        ptr_type,
        Some(layouts),
        Some(flat_count),
        Some(at),
    )
}

fn lower_flat_values_impl(
    memory: &crate::shared_memory::SharedMemory,
    realloc: &mut Realloc<'_>,
    max_flat: usize,
    values: &[Value],
    types: &[ValType],
    ptr_type: CoreType,
    layouts: Option<&FlatLayoutCache>,
    flat_count: Option<usize>,
    spill_at: Option<u32>,
) -> Result<Vec<CoreValue>, CanonError> {
    if values.len() != types.len() {
        return Err(CanonError::Unsupported(
            "lower_flat_values: value/type count mismatch",
        ));
    }
    let count = flat_count.unwrap_or_else(|| flat_types_count_bounded(types, max_flat));
    if count <= max_flat {
        let mut flat = Vec::with_capacity(count);
        for (v, t) in values.iter().zip(types) {
            lower_flat_into(memory, realloc, v, t, ptr_type, layouts, &mut flat)?;
        }
        return Ok(flat);
    }

    // Over the limit: store the arguments as a tuple in memory and pass the
    // pointer. The tuple is a record of the parameter types, so the layout is
    // the one `canon_value` already computes.
    let fallback_ty;
    let (tuple_ty, memory_layout) =
        if let Some((ty, layout)) = layouts.and_then(|cache| cache.tuple_for(max_flat)) {
            (ty, Some(layout))
        } else {
            fallback_ty = ValType::Record(
                types
                    .iter()
                    .enumerate()
                    .map(|(i, t)| (format!("{i}"), t.clone()))
                    .collect(),
            );
            (&fallback_ty, None)
        };
    let mut object = crate::value::Object::new();
    for (i, v) in values.iter().enumerate() {
        object.properties.insert(format!("{i}").into(), v.clone());
    }
    let tuple = Value::Object(crate::heap::alloc(object));

    let size = memory_layout.map_or_else(
        || crate::canon_layout::elem_size(tuple_ty),
        |layout| layout.size,
    );
    let align = memory_layout.map_or_else(
        || crate::canon_layout::alignment(tuple_ty),
        |layout| layout.align,
    );
    let at = match spill_at {
        Some(at) => at,
        None => realloc(size, align).ok_or(CanonError::OutOfMemory { size })?,
    };
    if let Some(layout) = memory_layout {
        crate::canon_value::store_with_layout(memory, realloc, &tuple, tuple_ty, at, layout)?;
    } else {
        crate::canon_value::store_with(memory, realloc, &tuple, tuple_ty, at)?;
    }
    Ok(if spill_at.is_some() {
        Vec::new()
    } else {
        vec![flat_from_ptr(at, ptr_type)]
    })
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
    lift_flat_values_impl(memory, max_flat, values, types, ptr_type, None, None)
}

pub(crate) fn lift_flat_values_cached(
    memory: &crate::shared_memory::SharedMemory,
    max_flat: usize,
    values: &[CoreValue],
    types: &[ValType],
    ptr_type: CoreType,
    layouts: &FlatLayoutCache,
    flat_count: usize,
) -> Result<Vec<Value>, CanonError> {
    lift_flat_values_impl(
        memory,
        max_flat,
        values,
        types,
        ptr_type,
        Some(layouts),
        Some(flat_count),
    )
}

fn lift_flat_values_impl(
    memory: &crate::shared_memory::SharedMemory,
    max_flat: usize,
    values: &[CoreValue],
    types: &[ValType],
    ptr_type: CoreType,
    layouts: Option<&FlatLayoutCache>,
    flat_count: Option<usize>,
) -> Result<Vec<Value>, CanonError> {
    if flat_count.unwrap_or_else(|| flat_types_count_bounded(types, max_flat)) <= max_flat {
        let mut vi = CoreValueIter::new(values);
        let mut out = Vec::with_capacity(types.len());
        for t in types {
            out.push(lift_flat_impl(memory, &mut vi, t, ptr_type, layouts)?);
        }
        return Ok(out);
    }

    let at = values
        .first()
        .copied()
        .ok_or(CanonError::Unsupported("lift_flat_values: no pointer"))?;
    let at = flat_ptr(at)?;
    let fallback_ty;
    let tuple = if let Some((ty, layout)) = layouts.and_then(|cache| cache.tuple_for(max_flat)) {
        crate::canon_value::load_with_layout(memory, ty, at, layout)?
    } else {
        fallback_ty = ValType::Record(
            types
                .iter()
                .enumerate()
                .map(|(i, t)| (format!("{i}"), t.clone()))
                .collect(),
        );
        crate::canon_value::load(memory, &fallback_ty, at)?
    };
    let mut out = Vec::with_capacity(types.len());
    for i in 0..types.len() {
        out.push(crate::canon_value::record_field_public(
            &tuple,
            &format!("{i}"),
        ));
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

    #[test]
    fn unsigned_scalars_round_trip_flat_without_sign_extension() {
        let memory = mem();
        let mut no_alloc = |_, _| None;
        let mut realloc: Realloc<'_> = &mut no_alloc;
        for bits in [0, 1 << 31, u32::MAX] {
            let value = Value::F64(bits as f64);
            let flat = lower_flat(&memory, &mut realloc, &value, &ValType::U32, CoreType::I32).unwrap();
            assert_eq!(flat, vec![CoreValue::I32(bits)]);
            assert!(matches!(round_trip(&value, &ValType::U32), Value::F64(n) if n == bits as f64));
        }
        for bits in [0, 1 << 63, u64::MAX] {
            let value = Value::bigint_u64(bits);
            let flat = lower_flat(&memory, &mut realloc, &value, &ValType::U64, CoreType::I32).unwrap();
            assert_eq!(flat, vec![CoreValue::I64(bits)]);
            assert!(matches!(round_trip(&value, &ValType::U64), Value::BigInt(n) if n.to_u64_wrapping() == bits && !n.is_negative()));
        }
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
        assert_eq!(
            crate::canon_value::record_field_public(&back, "a").as_i32(),
            3
        );
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
        small
            .properties
            .insert("tag".into(), Value::String("small".into()));
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
        wide.properties
            .insert("tag".into(), Value::String("wide".into()));
        wide.properties.insert("val".into(), Value::F64(2.5));
        let wide = Value::Object(crate::heap::alloc(wide));
        let back = round_trip(&wide, &t);
        assert_eq!(
            crate::canon_value::record_field_public(&back, "val").as_f64(),
            2.5,
            "an f64 case REINTERPRETED into an i64 slot must come back as 2.5, not its bit pattern"
        );
    }

    #[test]
    fn nested_variant_coerces_payload_in_place_without_shifting_next_field() {
        let t = ValType::Record(vec![
            (
                s("choice"),
                ValType::Variant(vec![
                    (
                        s("small"),
                        Some(ValType::Record(vec![
                            (s("integer"), ValType::I32),
                            (s("float"), ValType::F32),
                        ])),
                    ),
                    (
                        s("wide"),
                        Some(ValType::Record(vec![
                            (s("integer"), ValType::F64),
                            (s("float"), ValType::I64),
                        ])),
                    ),
                ]),
            ),
            (s("tail"), ValType::I32),
        ]);
        let mut payload = crate::value::Object::new();
        payload.properties.insert("integer".into(), Value::I32(42));
        payload.properties.insert("float".into(), Value::F64(2.5));
        let mut choice = crate::value::Object::new();
        choice
            .properties
            .insert("tag".into(), Value::String("small".into()));
        choice
            .properties
            .insert("val".into(), Value::Object(crate::heap::alloc(payload)));
        let mut record = crate::value::Object::new();
        record
            .properties
            .insert("choice".into(), Value::Object(crate::heap::alloc(choice)));
        record.properties.insert("tail".into(), Value::I32(9));
        let value = Value::Object(crate::heap::alloc(record));
        let memory = mem();
        let flat = with_alloc(&memory, |realloc| {
            lower_flat(&memory, realloc, &value, &t, CoreType::I32).unwrap()
        });
        assert_eq!(
            flat,
            vec![
                CoreValue::I32(0),
                CoreValue::I64(42),
                CoreValue::I64(2.5f32.to_bits() as u64),
                CoreValue::I32(9),
            ]
        );
        let mut values = CoreValueIter::new(&flat);
        let lifted = lift_flat(&memory, &mut values, &t, CoreType::I32).unwrap();
        assert!(values.done());
        assert_eq!(
            crate::canon_value::record_field_public(&lifted, "tail").as_i32(),
            9
        );
    }

    #[test]
    fn nested_lowering_preserves_errors_after_prefix_slots() {
        let memory = mem();
        let mut record = crate::value::Object::new();
        record.properties.insert("head".into(), Value::I32(4));
        record.properties.insert("items".into(), Value::Null);
        let value = Value::Object(crate::heap::alloc(record));
        let fixed = ValType::Record(vec![
            (s("head"), ValType::I32),
            (s("items"), ValType::ListFixed(Box::new(ValType::I32), 1)),
        ]);
        assert!(matches!(
            with_alloc(&memory, |realloc| lower_flat(
                &memory,
                realloc,
                &value,
                &fixed,
                CoreType::I32
            )),
            Err(CanonError::FixedListLength { got: 0, want: 1 })
        ));

        let unsupported =
            ValType::Record(vec![(s("head"), ValType::I32), (s("items"), ValType::Any)]);
        assert!(matches!(
            with_alloc(&memory, |realloc| lower_flat(
                &memory,
                realloc,
                &value,
                &unsupported,
                CoreType::I32
            )),
            Err(CanonError::Unsupported("any (not a component type)"))
        ));
    }

    #[test]
    fn cached_nested_layout_matches_uncached_flat_conversion() {
        let ft = std::sync::Arc::new(crate::canon_def::CanonFuncType {
            params: vec![ValType::Record(vec![
                (
                    s("choice"),
                    ValType::Variant(vec![(s("empty"), None), (s("number"), Some(ValType::F64))]),
                ),
                (s("optional"), ValType::Option(Box::new(ValType::I32))),
                (
                    s("result"),
                    ValType::Result(None, Some(Box::new(ValType::I64))),
                ),
            ])],
            result: None,
            is_async: false,
        });
        let layouts = FlatLayoutCache::for_functype(&ft, CoreType::I32);
        assert_eq!(layouts.variants.len(), 3);

        let mut error = crate::value::Object::new();
        error
            .properties
            .insert("__wasi_error".into(), Value::I64(19));
        let mut record = crate::value::Object::new();
        record
            .properties
            .insert("choice".into(), Value::String("empty".into()));
        record.properties.insert("optional".into(), Value::I32(7));
        record
            .properties
            .insert("result".into(), Value::Object(crate::heap::alloc(error)));
        let values = [Value::Object(crate::heap::alloc(record))];
        let memory = mem();
        let uncached = with_alloc(&memory, |realloc| {
            lower_flat_values(&memory, realloc, 16, &values, &ft.params, CoreType::I32).unwrap()
        });
        let cached = with_alloc(&memory, |realloc| {
            lower_flat_values_cached(
                &memory,
                realloc,
                16,
                &values,
                &ft.params,
                CoreType::I32,
                &layouts,
                layouts.params_flat_len,
            )
            .unwrap()
        });
        assert_eq!(cached, uncached);

        let back = lift_flat_values_cached(
            &memory,
            16,
            &cached,
            &ft.params,
            CoreType::I32,
            &layouts,
            layouts.params_flat_len,
        )
        .unwrap();
        let uncached_back =
            lift_flat_values(&memory, 16, &uncached, &ft.params, CoreType::I32).unwrap();
        for lifted in [&back[0], &uncached_back[0]] {
            assert_eq!(
                format!(
                    "{}",
                    crate::canon_value::record_field_public(lifted, "choice")
                ),
                "empty"
            );
            assert_eq!(
                crate::canon_value::record_field_public(lifted, "optional").as_i32(),
                7
            );
            let result = crate::canon_value::record_field_public(lifted, "result");
            assert_eq!(
                crate::canon_value::record_field_public(&result, "__wasi_error").as_i64(),
                19
            );
        }
    }

    /// Every case must occupy the same number of core values. A short case is
    /// PADDED on lower and DRAINED on lift; if either side forgets, the next
    /// parameter is read from this variant's tail.
    #[test]
    fn a_payload_free_case_still_occupies_the_joined_width() {
        let t = ValType::Variant(vec![(s("none"), None), (s("wide"), Some(ValType::F64))]);
        let m = mem();
        let bare = Value::String("none".into());
        let flat = with_alloc(&m, |r| lower_flat(&m, r, &bare, &t, CoreType::I32).unwrap());
        // discriminant + one joined payload slot, even though this case has no
        // payload at all.
        assert_eq!(flat.len(), 2, "a payload-free case must still be padded");

        let mut vi = CoreValueIter::new(&flat);
        let back = lift_flat(&m, &mut vi, &t, CoreType::I32).unwrap();
        assert!(
            vi.done(),
            "the unused slot must be drained, not left behind"
        );
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

    #[test]
    fn lowered_string_result_uses_the_callers_result_address() {
        let m = mem();
        let ft = crate::canon_def::CanonFuncType {
            params: vec![],
            result: Some(ValType::String),
            is_async: false,
        };
        let layouts = FlatLayoutCache::for_functype(&ft, CoreType::I32);
        let mut allocations = Vec::new();
        let mut allocate = |size, align| {
            allocations.push((size, align));
            Some(512)
        };
        let mut realloc: Realloc<'_> = &mut allocate;
        let flat = lower_flat_values_cached_into(
            &m,
            &mut realloc,
            crate::canon_flat::MAX_FLAT_RESULTS,
            &[Value::String("hi".into())],
            &[ValType::String],
            CoreType::I32,
            &layouts,
            layouts.result_flat_len,
            64,
        )
        .unwrap();
        assert!(flat.is_empty());
        let mut pair = [0; 8];
        m.read_bytes(64, &mut pair);
        assert_eq!(u32::from_le_bytes(pair[..4].try_into().unwrap()), 512);
        assert_eq!(u32::from_le_bytes(pair[4..].try_into().unwrap()), 2);
        let mut bytes = [0; 2];
        m.read_bytes(512, &mut bytes);
        assert_eq!(&bytes, b"hi");

        assert!(matches!(
            lower_flat_values_cached_into(
                &m,
                &mut realloc,
                crate::canon_flat::MAX_FLAT_RESULTS,
                &[Value::String("hi".into())],
                &[ValType::String],
                CoreType::I32,
                &layouts,
                layouts.result_flat_len,
                65,
            ),
            Err(crate::canon_value::CanonError::Misaligned { .. })
        ));
        assert_eq!(allocations, vec![(2, 1)]);
    }

    #[test]
    fn cached_spilled_tuple_matches_uncached_bytes_and_values() {
        let types = vec![ValType::I32; crate::canon_flat::MAX_FLAT_PARAMS + 2];
        let values: Vec<_> = (0..types.len()).map(|i| Value::I32(i as i32)).collect();
        let ft = crate::canon_def::CanonFuncType {
            params: types.clone(),
            result: None,
            is_async: false,
        };
        let layouts = FlatLayoutCache::for_functype(&ft, CoreType::I32);
        assert!(layouts.param_tuple.is_some());
        let plain = mem();
        let cached = mem();
        let plain_ptr = with_alloc(&plain, |r| {
            lower_flat_values(
                &plain,
                r,
                crate::canon_flat::MAX_FLAT_PARAMS,
                &values,
                &types,
                CoreType::I32,
            )
            .unwrap()
        });
        let cached_ptr = with_alloc(&cached, |r| {
            lower_flat_values_cached(
                &cached,
                r,
                crate::canon_flat::MAX_FLAT_PARAMS,
                &values,
                &types,
                CoreType::I32,
                &layouts,
                layouts.params_flat_len,
            )
            .unwrap()
        });
        assert_eq!(plain_ptr, cached_ptr);
        let mut a = [0; 128];
        let mut b = [0; 128];
        plain.read_bytes(512, &mut a);
        cached.read_bytes(512, &mut b);
        assert_eq!(a, b);
        let lifted = lift_flat_values_cached(
            &cached,
            crate::canon_flat::MAX_FLAT_PARAMS,
            &cached_ptr,
            &types,
            CoreType::I32,
            &layouts,
            layouts.params_flat_len,
        )
        .unwrap();
        for (i, value) in lifted.iter().enumerate() {
            assert_eq!(value.as_i32(), i as i32);
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

    #[test]
    fn cached_flat_lengths_preserve_abi_boundaries_without_expanding_fixed_lists() {
        let exact = crate::canon_def::CanonFuncType {
            params: vec![ValType::ListFixed(Box::new(ValType::U8), 16)],
            result: Some(ValType::I64),
            is_async: false,
        };
        let exact_cache = FlatLayoutCache::for_functype(&exact, CoreType::I32);
        assert_eq!(exact_cache.params_flat_len, 16);
        assert!(exact_cache.param_tuple.is_none());
        assert_eq!(exact_cache.result_flat_len, 1);
        assert!(exact_cache.result_tuple.is_none());

        let large = crate::canon_def::CanonFuncType {
            params: vec![ValType::ListFixed(Box::new(ValType::U8), 268_435_455)],
            result: Some(ValType::String),
            is_async: false,
        };
        let large_cache = FlatLayoutCache::for_functype(&large, CoreType::I32);
        assert_eq!(large_cache.params_flat_len, 17);
        assert!(large_cache.param_tuple.is_some());
        assert_eq!(large_cache.result_flat_len, 2);
        assert!(large_cache.result_tuple.is_some());
    }

    #[test]
    #[ignore = "paired allocation microbenchmark"]
    fn benchmark_bounded_flat_count_against_materialized_types() {
        use std::hint::black_box;
        use std::time::Instant;

        let types = vec![ValType::I32; 16];
        let rounds = 200_000;
        let before = Instant::now();
        let mut materialized = 0;
        for _ in 0..rounds {
            materialized += black_box(crate::canon_flat::flatten_types(
                black_box(&types),
                CoreType::I32,
            ))
            .len();
        }
        let materialized_time = before.elapsed();
        let before = Instant::now();
        let mut bounded = 0;
        for _ in 0..rounds {
            bounded += black_box(flat_types_count_bounded(
                black_box(types.iter()),
                crate::canon_flat::MAX_FLAT_PARAMS,
            ));
        }
        let bounded_time = before.elapsed();
        assert_eq!(bounded, materialized);
        eprintln!(
            "flat count, {rounds} calls: materialized={materialized_time:?}, bounded={bounded_time:?}"
        );
    }
}
