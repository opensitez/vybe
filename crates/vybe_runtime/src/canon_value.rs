//! Canonical ABI value transfer — `CanonicalABI.md` §`store` / §`load`.
//!
//! Moves ONE value of a component type between a `Value` and linear memory,
//! using the layout in [`crate::canon_layout`]. Together they are what
//! `canon future.{read,write}` needs: that built-in is
//! `(func (param i32 T) (result i32))` — a handle and a pointer, with **no
//! count** — because a future carries exactly one element whose size and
//! encoding come from its type.
//!
//! Both entry points begin with the two assertions the spec opens with:
//!
//! ```python
//! assert(ptr == align_to(ptr, alignment(t, ...)))
//! assert(ptr + elem_size(t, ...) <= len(cx.opts.memory))
//! ```
//!
//! Here they are real errors rather than debug assertions, because they are
//! reachable from guest code: a component that hands over a misaligned pointer
//! must be told, not silently allowed to write across a field boundary.
//!
//! **Not every `ValType` is covered**, and the uncovered ones return an error
//! naming themselves rather than moving an approximate number of bytes. A
//! canonical ABI that is quietly wrong about layout is worse than one that
//! refuses: the first corrupts a peer's memory, the second fails a test.

use crate::canon_layout::{PTR_SIZE, align_to, alignment, elem_size};
use crate::component::ValType;
use crate::value::Value;

/// Why a canonical transfer could not be performed.
#[derive(Debug, Clone)]
pub enum CanonError {
    Misaligned {
        ptr: u32,
        required: u32,
    },
    OutOfBounds {
        ptr: u32,
        size: u32,
        memory_len: usize,
    },
    /// The type has a canonical layout this crate has not implemented yet.
    /// Named so the message points at the work rather than at the symptom.
    Unsupported(&'static str),
    /// A flat core value did not have the type the ValType being lifted
    /// requires. Carries BOTH sides: this used to be an `Unsupported` whose
    /// helper did `let _ = (v, want);`, so the one fact a reader needs — what
    /// arrived versus what was wanted — was computed and then thrown away.
    FlatTypeMismatch {
        got: crate::canon_flat::CoreType,
        want: crate::canon_flat::CoreType,
    },
    /// `realloc` could not satisfy the allocation a `string`/`list` needed.
    OutOfMemory {
        size: u32,
    },
    /// 🔧 A fixed-length `list<T, N>` was given a different number of
    /// elements. `store_list` ASSERTS `N == len(v)`: storing fewer leaves the
    /// tail as stale memory that `load` reads back as data, storing more runs
    /// past the value into the next field.
    FixedListLength {
        got: usize,
        want: u32,
    },
    /// A `char` was not a Unicode scalar value.
    ///
    /// `CanonicalABI.md:2456 convert_i32_to_char` traps on `i >= 0x110000` and
    /// on the surrogate range `0xD800..=0xDFFF`. That check IS the difference
    /// between `char` and `u32`; without it a lone surrogate crosses the ABI
    /// as though it were a character.
    NotAScalarValue {
        got: u32,
    },
    /// A `variant` discriminant named a case the type does not have.
    ///
    /// This used to be an `Unsupported("variant discriminant out of range")`,
    /// which the Display impl renders as "no store/load implemented for …".
    /// That is a claim about THIS CRATE — it tells a reader the feature is
    /// missing when in fact the feature works and the PROGRAM is invalid.
    /// `enum`/`option`/`result` all despecialise through here, so it is the
    /// trap a wrong case index anywhere lands on.
    DiscriminantOutOfRange {
        got: u32,
        cases: usize,
    },
}

impl std::fmt::Display for CanonError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            CanonError::Misaligned { ptr, required } => write!(
                f,
                "canonical ABI: pointer {ptr} is not aligned to {required} bytes"
            ),
            CanonError::OutOfBounds {
                ptr,
                size,
                memory_len,
            } => write!(
                f,
                "canonical ABI: {size} bytes at {ptr} is out of bounds of linear memory ({memory_len} bytes)"
            ),
            CanonError::FlatTypeMismatch { got, want } => write!(
                f,
                "canonical ABI: flat lift wanted a core {want:?} and got a {got:?}"
            ),
            CanonError::Unsupported(what) => write!(
                f,
                "canonical ABI: no store/load implemented for {what} — see CanonicalABI.md §Storing/§Loading"
            ),
            CanonError::OutOfMemory { size } => write!(
                f,
                "canonical ABI: realloc could not supply {size} bytes for a string/list payload"
            ),
            CanonError::FixedListLength { got, want } => write!(
                f,
                "canonical ABI: a fixed-length `list<_, {want}>` was given {got} elements; \
                 the length is part of the TYPE and the elements are stored inline"
            ),
            CanonError::NotAScalarValue { got } => write!(
                f,
                "canonical ABI: {got:#x} is not a Unicode scalar value — a `char` must be \
                 below 0x110000 and outside the surrogate range 0xD800..=0xDFFF"
            ),
            CanonError::DiscriminantOutOfRange { got, cases } => write!(
                f,
                "canonical ABI: variant discriminant {got} names no case; the type has \
                 {cases} (0..{})",
                cases.saturating_sub(1)
            ),
        }
    }
}

/// The two guards every `store`/`load` opens with.
fn check(ptr: u32, t: &ValType, memory_len: usize) -> Result<u32, CanonError> {
    let required = alignment(t);
    if ptr != align_to(ptr, required) {
        return Err(CanonError::Misaligned { ptr, required });
    }
    let size = elem_size(t);
    if (ptr as usize).saturating_add(size as usize) > memory_len {
        return Err(CanonError::OutOfBounds {
            ptr,
            size,
            memory_len,
        });
    }
    Ok(size)
}

/// `cx.opts.realloc` — allocate `size` bytes at `align` and answer the address.
///
/// `CanonicalABI.md` threads this through every lift/lower because `string`
/// and `list` are (ptr, length) pairs: storing one means putting the ELEMENTS
/// somewhere first. `None` from the allocator is a genuine out-of-memory, not
/// a "not supported".
pub type Realloc<'a> = &'a mut dyn FnMut(u32, u32) -> Option<u32>;

/// An allocator that always fails, for the callers that store only scalars.
///
/// Kept explicit so a `string` reaching a no-realloc context reports
/// out-of-memory instead of silently writing a dangling (ptr, len).
fn no_realloc(_size: u32, _align: u32) -> Option<u32> {
    None
}

/// `store(cx, v, t, ptr)` — write one value of type `t` at `ptr`.
///
/// Scalar-only convenience: aggregates that need `realloc` report
/// `OutOfMemory`. Use [`store_with`] to supply one.
pub fn store(
    memory: &crate::shared_memory::SharedMemory,
    v: &Value,
    t: &ValType,
    ptr: u32,
) -> Result<(), CanonError> {
    let mut alloc = no_realloc;
    store_with(memory, &mut alloc, v, t, ptr)
}

/// `store(cx, v, t, ptr)` with the canonical option that makes aggregates
/// storable — `CanonicalABI.md` §`store`.
pub fn store_with(
    memory: &crate::shared_memory::SharedMemory,
    realloc: Realloc<'_>,
    v: &Value,
    t: &ValType,
    ptr: u32,
) -> Result<(), CanonError> {
    check(ptr, t, memory.len())?;
    let addr = ptr as usize;
    match t {
        // `store_int(cx, int(bool(v)), ptr, 1)` — a bool is one byte, 0 or 1,
        // never the raw truthiness of whatever the guest had.
        ValType::Bool => {
            let _ = memory.store_u8(addr, u8::from(v.as_bool()));
        }
        // The NARROW widths write exactly their own bytes. Writing four would
        // clobber whatever field the layout placed next.
        ValType::S8 | ValType::U8 => {
            let _ = memory.store_u8(addr, v.as_i32() as u8);
        }
        ValType::S16 | ValType::U16 => {
            let n = v.as_i32() as u16;
            let _ = memory.store_u8(addr, (n & 0xff) as u8);
            let _ = memory.store_u8(addr + 1, (n >> 8) as u8);
        }
        ValType::I32 => {
            let _ = memory.store_i32(addr, v.as_i32());
        }
        ValType::I64 => {
            let _ = memory.store_i64(addr, v.as_i64());
        }
        // `f32` is stored as its four-byte bit pattern, not as a truncated
        // `f64` — `store` is `store_int(encode_float_as_i32(v), ptr, 4)`.
        ValType::F32 => {
            let _ = memory.store_i32(addr, (v.as_f64() as f32).to_bits() as i32);
        }
        ValType::F64 => {
            let _ = memory.store_f64(addr, v.as_f64());
        }
        // A `char` is its scalar value in four bytes. It arrives as a
        // one-character string, which is what `load` produced.
        ValType::Char => {
            let _ = memory.store_i32(addr, value_to_scalar(v)? as i32);
        }
        // `store_flags` = `store_int(pack_flags_into_int(v, labels), ptr,
        // elem_size_flags(labels))` — the PACKED width, not four bytes.
        ValType::Flags(labels) => {
            let packed = pack_flags(v, labels);
            for k in 0..crate::canon_layout::flags_bytes(labels.len()) {
                let _ = memory.store_u8(addr + k as usize, (packed >> (8 * k)) as u8);
            }
        }
        // A handle is its i32 index into the handle table.
        ValType::Own(_)
        | ValType::Borrow(_)
        | ValType::Stream(_)
        | ValType::Future(_)
        | ValType::ErrorContext => {
            let _ = memory.store_i32(addr, v.as_i32());
        }
        // `store_string` — `CanonicalABI.md`: encode into freshly allocated
        // memory, then write the (ptr, length) pair. Length is BYTES of UTF-8,
        // never the string's code-unit count.
        ValType::String => {
            let text = format!("{v}");
            let bytes = text.as_bytes();
            let dest = realloc(bytes.len() as u32, 1).ok_or(CanonError::OutOfMemory {
                size: bytes.len() as u32,
            })?;
            write_bytes(memory, dest, bytes)?;
            let _ = memory.store_i32(addr, dest as i32);
            let _ = memory.store_i32(addr + PTR_SIZE as usize, bytes.len() as i32);
        }
        // `store_list` — same (ptr, length) shape, but length counts ELEMENTS
        // and each is stored at its own canonical stride.
        ValType::List(elem) => {
            let items = array_items(v);
            let stride = elem_size(elem);
            let dest = if items.is_empty() {
                0
            } else {
                let bytes = stride.saturating_mul(items.len() as u32);
                realloc(bytes, alignment(elem)).ok_or(CanonError::OutOfMemory { size: bytes })?
            };
            for (i, item) in items.iter().enumerate() {
                store_with(memory, realloc, item, elem, dest + stride * i as u32)?;
            }
            let _ = memory.store_i32(addr, dest as i32);
            let _ = memory.store_i32(addr + PTR_SIZE as usize, items.len() as i32);
        }
        // 🔧 `store_list` with a length present: `assert(N == len(v))` then
        // the elements INLINE. No `realloc`, no pointer — the list occupies
        // the value's own bytes.
        //
        // ⛔ The spec ASSERTS the count matches. Storing fewer would leave the
        // tail as whatever was in memory and `load` would read it back as
        // data; storing more would run past the value into the next field.
        ValType::ListFixed(elem, n) => {
            let items = array_items(v);
            if items.len() as u32 != *n {
                return Err(CanonError::FixedListLength {
                    got: items.len(),
                    want: *n,
                });
            }
            let stride = elem_size(elem);
            for (i, item) in items.iter().enumerate() {
                store_with(memory, realloc, item, elem, ptr + stride * i as u32)?;
            }
        }
        // `store_record` — each field aligned before it is placed, in
        // declaration order. Field lookup is by NAME, so a host value carrying
        // extra properties stores correctly and a missing one stores a default
        // rather than shifting every later field.
        ValType::Record(fields) => {
            let mut offset = 0u32;
            for (name, field_ty) in fields {
                offset = align_to(offset, alignment(field_ty));
                let field = record_field(v, name);
                store_with(memory, realloc, &field, field_ty, ptr + offset)?;
                offset += elem_size(field_ty);
            }
        }
        // `option` and `result` DESPECIALISE to `variant`, so they store
        // through the same path rather than growing their own copy of the
        // discriminant arithmetic.
        ValType::Option(inner) => {
            let (case, payload) = match v {
                Value::Null => (0u32, Value::Null),
                other => (1, other.clone()),
            };
            store_variant_parts(
                memory,
                realloc,
                ptr,
                case,
                &payload,
                &[("none".into(), None), ("some".into(), Some((**inner).clone()))],
            )?;
        }
        ValType::Result(ok, err) => {
            // An error carries `__wasi_error` in this tree; anything else is ok.
            //
            // The PAYLOAD of the error arm is that field's value, not the
            // wrapper object holding it. Passing the wrapper stored the error
            // as the string "[object]" — the round-trip test caught it
            // immediately, which a store-only test never would have.
            let carried = error_payload(v);
            let case = u32::from(carried.is_some());
            let payload = carried.unwrap_or_else(|| v.clone());
            store_variant_parts(
                memory,
                realloc,
                ptr,
                case,
                &payload,
                // A payload-free case stays `None`, so `store_variant_parts`
                // writes the discriminant and stops. `result<_, error-code>`
                // taking the ok path is exactly one byte.
                &[
                    ("ok".into(), ok.as_deref().cloned()),
                    ("error".into(), err.as_deref().cloned()),
                ],
            )?;
        }
        ValType::Variant(cases) => {
            let (case, payload) = variant_case(v, cases)?;
            store_variant_parts(memory, realloc, ptr, case, &payload, cases)?;
        }
        ValType::Any => return Err(CanonError::Unsupported("any (not a component type)")),
    }
    Ok(())
}

/// Read `len` bytes and decode them as UTF-8 — `CanonicalABI.md` §`load_string`.
///
/// Invalid UTF-8 is an ERROR, not a lossy substitution: the canonical ABI says
/// a `string` is well-formed UTF-8, so bytes that are not are a broken peer,
/// and replacing them with U+FFFD would hand the caller a string that never
/// existed.
///
/// Public because 📝 `error-context.new` is handed a raw `(ptr,
/// tagged_code_units)` pair rather than a `string`-typed value — the spec calls
/// this `load_string_from_range`.
pub fn read_utf8(
    memory: &crate::shared_memory::SharedMemory,
    ptr: usize,
    len: usize,
) -> Result<String, CanonError> {
    if ptr.saturating_add(len) > memory.len() {
        return Err(CanonError::OutOfBounds {
            ptr: ptr as u32,
            size: len as u32,
            memory_len: memory.len(),
        });
    }
    let mut bytes = vec![0u8; len];
    for (i, slot) in bytes.iter_mut().enumerate() {
        *slot = memory.load_u8(ptr + i).unwrap_or(0);
    }
    String::from_utf8(bytes).map_err(|_| CanonError::Unsupported("string is not valid UTF-8"))
}

/// The discriminant and payload of a stored `variant` — the read counterpart of
/// [`store_variant_parts`], so the two share one view of where a payload lives.
fn load_variant_parts(
    memory: &crate::shared_memory::SharedMemory,
    ptr: u32,
    cases: &[(String, Option<ValType>)],
) -> Result<(u32, Value), CanonError> {
    let addr = ptr as usize;
    let case = match crate::canon_layout::variant_discriminant_size(cases) {
        1 => memory.load_u8(addr).unwrap_or(0) as u32,
        2 => {
            let low = memory.load_u8(addr).unwrap_or(0) as u32;
            let high = memory.load_u8(addr + 1).unwrap_or(0) as u32;
            low | (high << 8)
        }
        _ => memory.load_i32(addr).unwrap_or(0) as u32,
    };
    let payload = match cases.get(case as usize) {
        Some((_, Some(payload_ty))) => {
            let offset = crate::canon_layout::variant_payload_offset(cases);
            load(memory, payload_ty, ptr + offset)?
        }
        _ => Value::Null,
    };
    Ok((case, payload))
}

/// Copy raw bytes into linear memory one at a time.
///
/// `SharedMemory` exposes typed stores only; a `string`'s payload is a byte
/// run with no alignment requirement, so this is the honest primitive rather
/// than pretending it is a sequence of i32s.
fn write_bytes(
    memory: &crate::shared_memory::SharedMemory,
    ptr: u32,
    bytes: &[u8],
) -> Result<(), CanonError> {
    let end = (ptr as usize).saturating_add(bytes.len());
    if end > memory.len() {
        return Err(CanonError::OutOfBounds {
            ptr,
            size: bytes.len() as u32,
            memory_len: memory.len(),
        });
    }
    for (i, byte) in bytes.iter().enumerate() {
        let _ = memory.store_u8(ptr as usize + i, *byte);
    }
    Ok(())
}

/// The elements of an array-shaped `Value`, or empty when it is not one.
/// The elements of an array-shaped value, for callers outside this module.
///
/// `canon_flat_values` needs the same view `store` uses — reading it a second
/// way is how a fixed list's flat lowering and its memory store would come to
/// disagree about what counts as an element.
pub fn array_items_public(v: &Value) -> Vec<Value> {
    array_items(v)
}

fn array_items(v: &Value) -> Vec<Value> {
    use crate::value::ObjectKind;
    if let Value::Object(object) = v {
        let object = object.lock().unwrap();
        if let ObjectKind::Array(items) = &object.kind {
            return items.clone();
        }
    }
    Vec::new()
}

/// One named field of a record-shaped `Value`.
///
/// A missing field stores as `Null`, which each type's arm turns into its own
/// zero. That keeps every LATER field at the right offset — the alternative,
/// erroring out, would leave a half-written record in the peer's memory.
fn record_field(v: &Value, name: &str) -> Value {
    if let Value::Object(object) = v {
        let object = object.lock().unwrap();
        if let Some(found) = object.properties.get(name) {
            return found.clone();
        }
        // WIT spells some fields `%type`; this tree stores them unprefixed.
        if let Some(found) = object.properties.get(name.trim_start_matches('%')) {
            return found.clone();
        }
    }
    Value::Null
}

/// The error payload of a `result`, if this value carries one.
fn error_payload(v: &Value) -> Option<Value> {
    if let Value::Object(object) = v {
        let object = object.lock().unwrap();
        return object.properties.get("__wasi_error").cloned();
    }
    None
}

/// `store_variant(cx, v, ptr, cases)` — the discriminant, then the payload at
/// `align_to(discriminant_size, max_case_alignment)`.
fn store_variant_parts(
    memory: &crate::shared_memory::SharedMemory,
    realloc: Realloc<'_>,
    ptr: u32,
    case: u32,
    payload: &Value,
    cases: &[(String, Option<ValType>)],
) -> Result<(), CanonError> {
    let addr = ptr as usize;
    match crate::canon_layout::variant_discriminant_size(cases) {
        1 => {
            let _ = memory.store_u8(addr, case as u8);
        }
        2 => {
            // No `store_u16`; the two bytes are little-endian like every other
            // canonical integer.
            let _ = memory.store_u8(addr, (case & 0xff) as u8);
            let _ = memory.store_u8(addr + 1, ((case >> 8) & 0xff) as u8);
        }
        _ => {
            let _ = memory.store_i32(addr, case as i32);
        }
    }
    if let Some(Some(payload_ty)) = cases.get(case as usize).map(|(_, t)| t) {
        let offset = crate::canon_layout::variant_payload_offset(cases);
        store_with(memory, realloc, payload, payload_ty, ptr + offset)?;
    }
    Ok(())
}

/// Which case of `cases` this value is, and the payload to store with it.
///
/// A payload-free case is named by its STRING — that is how this tree's hosts
/// spell `descriptor-type` (`"regular-file"`) and it is what an `enum` looks
/// like once despecialised. A case with a payload is `{ tag, val }`.
fn variant_case(
    v: &Value,
    cases: &[(String, Option<ValType>)],
) -> Result<(u32, Value), CanonError> {
    let named = |name: &str| cases.iter().position(|(case, _)| case == name);

    if let Value::String(text) = v {
        if let Some(index) = named(text) {
            return Ok((index as u32, Value::Null));
        }
    }
    if let Value::Object(object) = v {
        let object = object.lock().unwrap();
        if let Some(Value::String(tag)) = object.properties.get("tag") {
            if let Some(index) = named(tag) {
                let payload = object.properties.get("val").cloned().unwrap_or(Value::Null);
                return Ok((index as u32, payload));
            }
        }
    }
    Err(CanonError::Unsupported("variant case not named by the value"))
}

/// `load(cx, ptr, t)` — read one value of type `t` from `ptr`.
pub fn load(
    memory: &crate::shared_memory::SharedMemory,
    t: &ValType,
    ptr: u32,
) -> Result<Value, CanonError> {
    check(ptr, t, memory.len())?;
    let addr = ptr as usize;
    Ok(match t {
        // `convert_int_to_bool` — any non-zero byte is true.
        ValType::Bool => Value::Bool(memory.load_u8(addr).unwrap_or(0) != 0),
        // ⛔ SIGNED widths SIGN-EXTEND; unsigned ones do not. `CanonicalABI.md:2376`
        // spells them as separate cases — `load_int(cx, ptr, 1, signed = True)`
        // versus `load_int(cx, ptr, 1)` — so an `s8` holding `0xFF` lifts as
        // `-1` while a `u8` holding the same byte lifts as `255`.
        ValType::S8 => Value::I32(memory.load_u8(addr).unwrap_or(0) as i8 as i32),
        ValType::U8 => Value::I32(memory.load_u8(addr).unwrap_or(0) as i32),
        ValType::S16 => Value::I32(load_u16(memory, addr) as i16 as i32),
        ValType::U16 => Value::I32(load_u16(memory, addr) as i32),
        ValType::I32 => Value::I32(memory.load_i32(addr).unwrap_or(0)),
        ValType::I64 => Value::I64(memory.load_i64(addr).unwrap_or(0)),
        ValType::F32 => {
            Value::F64(f32::from_bits(memory.load_i32(addr).unwrap_or(0) as u32) as f64)
        }
        ValType::F64 => Value::F64(memory.load_f64(addr).unwrap_or(0.0)),
        ValType::Char => {
            let raw = memory.load_i32(addr).unwrap_or(0) as u32;
            Value::String(std::sync::Arc::from(scalar_to_char(raw)?.to_string().as_str()))
        }
        // `load_flags` = `unpack_flags_from_int(load_int(ptr, elem_size))` —
        // read only the packed width, then one bit per label.
        ValType::Flags(labels) => {
            let mut packed: u32 = 0;
            for k in 0..crate::canon_layout::flags_bytes(labels.len()) {
                packed |= (memory.load_u8(addr + k as usize).unwrap_or(0) as u32) << (8 * k);
            }
            unpack_flags(packed, labels)
        }
        ValType::Own(_)
        | ValType::Borrow(_)
        | ValType::Stream(_)
        | ValType::Future(_)
        | ValType::ErrorContext => {
            Value::I32(memory.load_i32(addr).unwrap_or(0))
        }
        // `load_string` — (ptr, length) where length is BYTES of UTF-8.
        ValType::String => {
            let at = memory.load_i32(addr).unwrap_or(0) as usize;
            let len = memory.load_i32(addr + PTR_SIZE as usize).unwrap_or(0) as usize;
            Value::String(std::sync::Arc::from(read_utf8(memory, at, len)?.as_str()))
        }
        // `load_list` — (ptr, length) where length counts ELEMENTS.
        ValType::List(elem) => {
            let at = memory.load_i32(addr).unwrap_or(0) as u32;
            let count = memory.load_i32(addr + PTR_SIZE as usize).unwrap_or(0) as u32;
            let stride = elem_size(elem);
            let mut items = Vec::with_capacity(count as usize);
            for i in 0..count {
                items.push(load(memory, elem, at + stride * i)?);
            }
            Value::Object(crate::heap::alloc(crate::value::Object::new_array(items)))
        }
        // 🔧 `load_list_from_valid_range(cx, ptr, N, t)` — N elements read
        // in place. The count comes from the TYPE, not from memory.
        ValType::ListFixed(elem, n) => {
            let stride = elem_size(elem);
            let mut items = Vec::with_capacity(*n as usize);
            for i in 0..*n {
                items.push(load(memory, elem, ptr + stride * i)?);
            }
            Value::Object(crate::heap::alloc(crate::value::Object::new_array(items)))
        }
        // `load_record` — fields in declaration order, each at its aligned offset.
        ValType::Record(fields) => {
            let mut object = crate::value::Object::new();
            let mut offset = 0u32;
            for (name, field_ty) in fields {
                offset = align_to(offset, alignment(field_ty));
                let value = load(memory, field_ty, ptr + offset)?;
                object.properties.insert(name.as_str().into(), value);
                offset += elem_size(field_ty);
            }
            Value::Object(crate::heap::alloc(object))
        }
        // `option` and `result` despecialise to `variant`, so they load through
        // the same discriminant arithmetic rather than repeating it.
        ValType::Option(inner) => {
            let cases: [(String, Option<ValType>); 2] =
                [("none".into(), None), ("some".into(), Some((**inner).clone()))];
            let (case, payload) = load_variant_parts(memory, ptr, &cases)?;
            // `none` is `Value::Null`, which is how every host here spells it.
            if case == 0 { Value::Null } else { payload }
        }
        ValType::Result(ok, err) => {
            let cases: [(String, Option<ValType>); 2] = [
                ("ok".into(), ok.as_deref().cloned()),
                ("error".into(), err.as_deref().cloned()),
            ];
            let (case, payload) = load_variant_parts(memory, ptr, &cases)?;
            if case == 0 {
                payload
            } else {
                // An error is spelled `{__wasi_error: …}` throughout this tree,
                // so a loaded error arrives in the shape `store` would accept
                // back — the two directions have to agree or a round trip
                // silently changes an error into an ok.
                let mut object = crate::value::Object::new();
                object.properties.insert("__wasi_error".into(), payload);
                Value::Object(crate::heap::alloc(object))
            }
        }
        // `load_variant` — read the discriminant, name the case, and load the
        // payload if that case carries one.
        ValType::Variant(cases) => {
            let (case, payload) = load_variant_parts(memory, ptr, cases)?;
            match cases.get(case as usize) {
                Some((name, None)) => Value::String(std::sync::Arc::from(name.as_str())),
                // A case that DECLARES a payload but did not carry one loads
                // as the bare name too, not as `{ tag, val: null }`.
                //
                // The split is on the VALUE, not on the declared type, because
                // every producer in this tree spells a payload-free case as a
                // bare string — `descriptor-type`'s `other(option<string>)`
                // arrives from the host as `"other"`. Keying on the type would
                // make that one case round-trip into a shape no consumer reads:
                // `fs_path::emit_type_is` and python's `emit_entry_type_is`
                // both compare `type` against a string, and an object compares
                // equal to nothing. `Option` already resolves this the same way
                // by loading `none` as `Value::Null` rather than a tag.
                Some((name, Some(_))) if matches!(payload, Value::Null) => {
                    Value::String(std::sync::Arc::from(name.as_str()))
                }
                Some((name, Some(_))) => {
                    // A payload-carrying case is `{ tag, val }` — the shape
                    // `variant_case` reads back on the store side.
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
        ValType::Any => return Err(CanonError::Unsupported("any (not a component type)")),
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::shared_memory::SharedMemory;

    fn mem() -> SharedMemory {
        SharedMemory::new(1024)
    }

    #[test]
    fn scalars_round_trip() {
        let m = mem();
        store(&m, &Value::I32(0x1234_5678), &ValType::I32, 0).unwrap();
        assert_eq!(load(&m, &ValType::I32, 0).unwrap().as_i32(), 0x1234_5678);

        store(&m, &Value::I64(-9), &ValType::I64, 8).unwrap();
        assert_eq!(load(&m, &ValType::I64, 8).unwrap().as_i64(), -9);

        store(&m, &Value::F64(1.5), &ValType::F64, 16).unwrap();
        assert_eq!(load(&m, &ValType::F64, 16).unwrap().as_f64(), 1.5);
    }

    #[test]
    fn bool_stores_one_byte_not_truthiness() {
        let m = mem();
        store(&m, &Value::Bool(true), &ValType::Bool, 0).unwrap();
        assert_eq!(m.load_u8(0).unwrap(), 1);
        // and any non-zero byte reads back as true
        m.store_u8(1, 0x7f).unwrap();
        assert!(load(&m, &ValType::Bool, 1).unwrap().as_bool());
    }

    #[test]
    fn misaligned_pointer_is_an_error_not_a_silent_write() {
        let m = mem();
        // i64 needs 8-byte alignment; 4 is not.
        let e = store(&m, &Value::I64(1), &ValType::I64, 4).unwrap_err();
        assert!(matches!(e, CanonError::Misaligned { required: 8, .. }));
    }

    #[test]
    fn out_of_bounds_is_an_error() {
        let m = SharedMemory::new(16);
        let e = store(&m, &Value::I64(1), &ValType::I64, 16).unwrap_err();
        assert!(matches!(e, CanonError::OutOfBounds { .. }));
    }

    #[test]
    fn unimplemented_types_refuse_rather_than_guess() {
        let m = mem();
        // The point: a wrong number of bytes would corrupt a peer's memory
        // silently. Refusing fails a test instead.
        //
        // This used to name `String`, which is now IMPLEMENTED — the test was
        // pinning a limitation that had been lifted, exactly like the `Any`
        // stand-ins in `clock.rs`/`io.rs` and this crate's own
        // "as this crate can spell it" comment. `Any` is the honest subject
        // now: it is not a component type at all, so it has no canonical
        // layout to be right or wrong about.
        assert!(matches!(
            load(&m, &ValType::Any, 0),
            Err(CanonError::Unsupported(_))
        ));
        assert!(matches!(
            store(&m, &Value::I32(1), &ValType::Any, 0),
            Err(CanonError::Unsupported(_))
        ));
        // And a `string` with no realloc reports OUT OF MEMORY rather than
        // writing a dangling (ptr, len) that points at nothing.
        assert!(matches!(
            store(&m, &Value::String("hi".into()), &ValType::String, 0),
            Err(CanonError::OutOfMemory { .. })
        ));
    }

    /// `store` then `load` must answer the value that went in — for the
    /// aggregates, not just the scalars.
    ///
    /// The two directions have to agree on shape or a round trip silently
    /// changes meaning: a `result`'s error arm re-read as an ok is the case
    /// that motivated pinning this.
    #[test]
    fn aggregates_round_trip_through_linear_memory() {
        let m = mem();
        // A bump allocator over the tail of the test's memory.
        let mut next = 256u32;
        let mut alloc = |size: u32, align: u32| -> Option<u32> {
            let at = crate::canon_layout::align_to(next, align.max(1));
            next = at + size;
            (next as usize <= m.len()).then_some(at)
        };

        let string_t = ValType::String;
        store_with(&m, &mut alloc, &Value::String("héllo".into()), &string_t, 0).unwrap();
        assert_eq!(format!("{}", load(&m, &string_t, 0).unwrap()), "héllo");

        // `result<_, error-code>` taking the ERROR arm — the shape every WASI
        // 0.3.1 call returns.
        let result_t = ValType::Result(None, Some(Box::new(ValType::String)));
        let mut errored = crate::value::Object::new();
        errored
            .properties
            .insert("__wasi_error".into(), Value::String("no-entry".into()));
        let errored = Value::Object(crate::heap::alloc(errored));
        store_with(&m, &mut alloc, &errored, &result_t, 16).unwrap();
        let back = load(&m, &result_t, 16).unwrap();
        let Value::Object(object) = &back else {
            panic!("an error arm must load back as an error object, got {back:?}");
        };
        assert_eq!(
            object
                .lock()
                .unwrap()
                .properties
                .get("__wasi_error")
                .map(|v| format!("{v}")),
            Some("no-entry".to_string()),
            "the error arm must not round-trip into an ok"
        );
    }

    /// `wasi:filesystem`'s `directory-entry`, including the ONE case that
    /// carries a payload.
    ///
    /// `descriptor-type`'s eight cases are payload-free except `other(option
    /// <string>)`, and that case is the reason the record is 24 bytes rather
    /// than the 12 an enum tag would give — it sets the variant's size and
    /// therefore where `name` lands. Every case except `other` can be stored
    /// correctly by code that has the payload offset wrong, because there is no
    /// payload to put there.
    ///
    /// It is also the case a real filesystem produces least often, so it is the
    /// one a test has to reach deliberately. The host answered an undeclared
    /// `"unknown"` here for months, and the symptom was not a mislabelled entry
    /// — `variant_case` refused, the copy failed, and one fifo made every entry
    /// beside it unreadable.
    #[test]
    fn a_directory_entry_round_trips_including_the_payload_carrying_case() {
        let m = SharedMemory::new(4096);
        let mut next = 512u32;
        let mut alloc = |size: u32, align: u32| -> Option<u32> {
            let at = crate::canon_layout::align_to(next, align.max(1));
            next = at + size;
            (next as usize <= m.len()).then_some(at)
        };

        let case = |name: &str| (name.to_string(), None);
        let descriptor_type = ValType::Variant(vec![
            case("block-device"),
            case("character-device"),
            case("directory"),
            case("fifo"),
            case("symbolic-link"),
            case("regular-file"),
            case("socket"),
            (
                "other".to_string(),
                Some(ValType::Option(Box::new(ValType::String))),
            ),
        ]);
        let entry_t = ValType::Record(vec![
            ("type".to_string(), descriptor_type),
            ("name".to_string(), ValType::String),
        ]);

        // The layout the guest lift in `fs_path::emit_read_directory_entries`
        // reads by. Pinned because that side computes the same numbers from
        // `canon_layout` independently — a WIT change that moves them must
        // break BOTH, not silently desynchronise them.
        assert_eq!(crate::canon_layout::elem_size(&entry_t), 24);
        assert_eq!(crate::canon_layout::alignment(&entry_t), 4);

        let entry = |kind: &str, name: &str| {
            let mut o = crate::value::Object::new();
            o.properties.insert("type".into(), Value::String(kind.into()));
            o.properties.insert("name".into(), Value::String(name.into()));
            Value::Object(crate::heap::alloc(o))
        };
        let field = |v: &Value, key: &str| {
            let Value::Object(o) = v else {
                panic!("expected a directory-entry record, got {v:?}");
            };
            let o = o.lock().unwrap();
            format!("{}", o.properties.get(key).cloned().unwrap_or(Value::Null))
        };

        // Two entries back to back at the canonical stride, as a stream copy
        // lays them out — the second one is where a wrong stride shows up.
        for (i, (kind, name)) in [("regular-file", "a.txt"), ("other", "pipe")]
            .into_iter()
            .enumerate()
        {
            let at = 24 * i as u32;
            store_with(&m, &mut alloc, &entry(kind, name), &entry_t, at).unwrap();
            let back = load(&m, &entry_t, at).unwrap();
            assert_eq!(field(&back, "type"), kind, "case name must survive");
            assert_eq!(field(&back, "name"), name, "entry name must survive");
        }

        // A case the variant does not declare must REFUSE. Storing it as some
        // near-miss index would relabel the entry, and there is no index that
        // means "not one of these".
        assert!(matches!(
            store_with(&m, &mut alloc, &entry("unknown", "x"), &entry_t, 48),
            Err(CanonError::Unsupported(_))
        ));

        // `other` CARRYING a description keeps the `{ tag, val }` shape — the
        // bare-name form above is what an absent payload loads as, and the two
        // must stay distinguishable or the description is silently dropped.
        let mut described = crate::value::Object::new();
        described
            .properties
            .insert("tag".into(), Value::String("other".into()));
        described
            .properties
            .insert("val".into(), Value::String("door".into()));
        let described = Value::Object(crate::heap::alloc(described));
        store_with(&m, &mut alloc, &described, &entry_t.clone(), 72).unwrap_err();

        let other_t = ValType::Variant(vec![(
            "other".to_string(),
            Some(ValType::Option(Box::new(ValType::String))),
        )]);
        store_with(&m, &mut alloc, &described, &other_t, 96).unwrap();
        let back = load(&m, &other_t, 96).unwrap();
        assert_eq!(field(&back, "tag"), "other");
        assert_eq!(field(&back, "val"), "door");
    }
}

// ── Shared with the FLAT representation ───────────────────────────────────
//
// `canon_flat_values` needs the same value-shape decisions this module already
// makes — which property is a record field, how a `result` splits into
// (case, payload), how a variant names its case. Re-deriving them there would
// be two answers to one question, and the memory and register forms of the
// same value would disagree about which case is live.
//
// These are thin `pub` re-exports of the private helpers above, NOT new logic.

/// A record field by name, tolerating the WIT `%`-prefix spelling.
pub fn record_field_public(v: &Value, name: &str) -> Value {
    record_field(v, name)
}

/// A `result` as `(case_index, payload)` — 0 = ok, 1 = error.
///
/// Errors are spelled `{__wasi_error: …}` throughout this tree, which is what
/// makes the split total: anything without that key is an `ok`.
pub fn result_parts_public(v: &Value) -> (u32, Value) {
    match error_payload(v) {
        Some(payload) => (1, payload),
        None => (0, v.clone()),
    }
}

/// A variant value as `(case_index, payload)`.
pub fn variant_case_public(
    v: &Value,
    cases: &[(String, Option<ValType>)],
) -> Result<(u32, Value), CanonError> {
    variant_case(v, cases)
}

/// Decode a (ptr, length) pair that arrived FLAT rather than from memory.
///
/// The bytes are in linear memory either way — only the pair's location
/// differs — so this writes the pair into a scratch cell and reuses `load`,
/// which is the spec's own "reuse the previous definitions" for strings and
/// lists.
pub fn load_pair_public(
    memory: &crate::shared_memory::SharedMemory,
    t: &ValType,
    at: u32,
    len: u32,
) -> Result<Value, CanonError> {
    match t {
        ValType::String => {
            let text = read_utf8(memory, at as usize, len as usize)?;
            Ok(Value::String(std::sync::Arc::from(text.as_str())))
        }
        ValType::List(elem) => {
            let stride = elem_size(elem);
            let mut items = Vec::with_capacity(len as usize);
            for i in 0..len {
                items.push(load(memory, elem, at + stride * i)?);
            }
            Ok(Value::Object(crate::heap::alloc(
                crate::value::Object::new_array(items),
            )))
        }
        _ => Err(CanonError::Unsupported("load_pair: not a (ptr, length) type")),
    }
}

/// Store a string or list and answer its `(ptr, length)` pair, for a caller
/// that will pass the pair in core values instead of writing it to memory.
pub fn store_pair_public(
    memory: &crate::shared_memory::SharedMemory,
    realloc: &mut Realloc<'_>,
    v: &Value,
    t: &ValType,
) -> Result<(u32, u32), CanonError> {
    match t {
        ValType::String => {
            let text = format!("{v}");
            let bytes = text.as_bytes();
            let size = bytes.len() as u32;
            let at = realloc(size.max(1), 1).ok_or(CanonError::OutOfMemory { size })?;
            write_bytes(memory, at, bytes)?;
            Ok((at, size))
        }
        ValType::List(elem) => {
            let items = array_items(v);
            let stride = elem_size(elem);
            let align = alignment(elem);
            let size = stride * items.len() as u32;
            let at = realloc(size.max(1), align).ok_or(CanonError::OutOfMemory { size })?;
            for (i, item) in items.iter().enumerate() {
                store_with(memory, realloc, item, elem, at + stride * i as u32)?;
            }
            Ok((at, items.len() as u32))
        }
        _ => Err(CanonError::Unsupported("store_pair: not a (ptr, length) type")),
    }
}

/// Two bytes, little-endian — the memory has no 16-bit primitive.
fn load_u16(memory: &crate::shared_memory::SharedMemory, addr: usize) -> u16 {
    let lo = memory.load_u8(addr).unwrap_or(0) as u16;
    let hi = memory.load_u8(addr + 1).unwrap_or(0) as u16;
    lo | (hi << 8)
}

/// `convert_i32_to_char` — `CanonicalABI.md:2456`.
///
/// ⛔ This validation is not optional and not a nicety: it is the ONLY thing
/// separating `char` from `u32`. Both are four bytes and both flatten to one
/// core `i32`; the trap is the difference.
pub fn scalar_to_char(raw: u32) -> Result<char, CanonError> {
    if raw >= 0x110000 || (0xD800..=0xDFFF).contains(&raw) {
        return Err(CanonError::NotAScalarValue { got: raw });
    }
    char::from_u32(raw).ok_or(CanonError::NotAScalarValue { got: raw })
}

/// The scalar value of a component `char`, which travels as a one-character
/// string. An empty or multi-character string is a program error, not a
/// representation this can guess at.
pub fn value_to_scalar(v: &Value) -> Result<u32, CanonError> {
    if let Value::String(s) = v {
        let mut it = s.chars();
        if let (Some(c), None) = (it.next(), it.next()) {
            return Ok(c as u32);
        }
    }
    // An i32 is accepted so a lowered `char` can round-trip through a numeric
    // slot, but it is still validated.
    let raw = v.as_i32() as u32;
    scalar_to_char(raw).map(|c| c as u32)
}

/// `pack_flags_into_int` — `CanonicalABI.md:3062`. Bit `k` is `labels[k]`.
///
/// ⛔ POSITIONAL: the same integer means different things under two different
/// label lists, which is why the labels travel with the type rather than being
/// looked up somewhere. A missing property is `false`, matching `int(bool(v[l]))`
/// on a record the guest filled partially.
pub fn pack_flags(v: &Value, labels: &[String]) -> u32 {
    let mut packed = 0u32;
    for (shift, label) in labels.iter().enumerate() {
        let set = match v {
            Value::Object(_) => record_field_public(v, label).as_bool(),
            // A flags value lowered from a plain integer keeps its bits — this
            // is how a round trip through a numeric slot survives.
            _ => (v.as_i32() as u32) >> shift & 1 == 1,
        };
        if set {
            packed |= 1 << shift;
        }
    }
    packed
}

/// `unpack_flags_from_int` — `CanonicalABI.md:3064`. A record of label → bool.
pub fn unpack_flags(packed: u32, labels: &[String]) -> Value {
    let mut object = crate::value::Object::new();
    for (shift, label) in labels.iter().enumerate() {
        object
            .properties
            .insert(label.as_str().into(), Value::Bool((packed >> shift) & 1 == 1));
    }
    Value::Object(crate::heap::alloc(object))
}
