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

use crate::canon_layout::{align_to, alignment, elem_size};
#[cfg(test)]
use crate::canon_layout::PTR_SIZE;
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
    ListByteLengthExceeded {
        bytes: usize,
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
    UnsignedOutOfRange {
        bits: u8,
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
            CanonError::UnsignedOutOfRange { bits } => write!(
                f,
                "canonical ABI: value is not an unsigned {bits}-bit integer"
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
            CanonError::ListByteLengthExceeded { bytes } => write!(
                f,
                "canonical ABI: list payload of {bytes} bytes exceeds MAX_LIST_BYTE_LENGTH"
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

pub(crate) fn unsigned_u32(v: &Value) -> Result<u32, CanonError> {
    match v {
        Value::I32(n) if *n >= 0 => Ok(*n as u32),
        Value::I64(n) if *n >= 0 && *n <= u32::MAX as i64 => Ok(*n as u32),
        Value::F64(n) if n.is_finite() && *n >= 0.0 && *n < 4_294_967_296.0 && n.fract() == 0.0 => Ok(*n as u32),
        Value::F32(n) if n.is_finite() && *n >= 0.0 && (*n as f64) < 4_294_967_296.0 && n.fract() == 0.0 => Ok(*n as u32),
        Value::BigInt(n) if n.fits_u32() => Ok(n.to_u64_wrapping() as u32),
        _ => Err(CanonError::UnsignedOutOfRange { bits: 32 }),
    }
}

pub(crate) fn unsigned_u64(v: &Value) -> Result<u64, CanonError> {
    match v {
        Value::I32(n) if *n >= 0 => Ok(*n as u64),
        Value::I64(n) if *n >= 0 => Ok(*n as u64),
        Value::F64(n) if n.is_finite() && *n >= 0.0 && *n < 18_446_744_073_709_551_616.0 && n.fract() == 0.0 => Ok(*n as u64),
        Value::F32(n) if n.is_finite() && *n >= 0.0 && (*n as f64) < 18_446_744_073_709_551_616.0 && n.fract() == 0.0 => Ok(*n as u64),
        Value::BigInt(n) if !n.is_negative() && n.bit_len() <= 64 => Ok(n.to_u64_wrapping()),
        _ => Err(CanonError::UnsignedOutOfRange { bits: 64 }),
    }
}

/// The two guards every `store`/`load` opens with.
fn check(ptr: u32, t: &ValType, memory_len: usize) -> Result<u32, CanonError> {
    check_layout(ptr, alignment(t), elem_size(t), memory_len)
}

fn check_layout(ptr: u32, required: u32, size: u32, memory_len: usize) -> Result<u32, CanonError> {
    if ptr != align_to(ptr, required) {
        return Err(CanonError::Misaligned { ptr, required });
    }
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

fn string_text(v: &Value) -> std::borrow::Cow<'_, str> {
    match v {
        Value::String(text) => std::borrow::Cow::Borrowed(text.as_ref()),
        _ => std::borrow::Cow::Owned(format!("{v}")),
    }
}

fn store_pair_header(memory: &crate::shared_memory::SharedMemory, addr: usize, ptr: u32, len: u32) {
    let mut bytes = [0u8; 8];
    bytes[..4].copy_from_slice(&ptr.to_le_bytes());
    bytes[4..].copy_from_slice(&len.to_le_bytes());
    let _ = memory.write_array(addr, &bytes);
}

fn load_pair_header(memory: &crate::shared_memory::SharedMemory, addr: usize) -> (u32, u32) {
    let bytes = memory.read_array::<8>(addr).unwrap_or([0; 8]);
    (
        u32::from_le_bytes(bytes[..4].try_into().unwrap()),
        u32::from_le_bytes(bytes[4..].try_into().unwrap()),
    )
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
    store_impl(memory, realloc, v, t, ptr, None)
}

pub(crate) fn store_with_layout(
    memory: &crate::shared_memory::SharedMemory,
    realloc: Realloc<'_>,
    v: &Value,
    t: &ValType,
    ptr: u32,
    layout: &crate::canon_layout::MemLayout,
) -> Result<(), CanonError> {
    store_impl(memory, realloc, v, t, ptr, Some(layout))
}

fn store_impl(
    memory: &crate::shared_memory::SharedMemory,
    realloc: Realloc<'_>,
    v: &Value,
    t: &ValType,
    ptr: u32,
    layout: Option<&crate::canon_layout::MemLayout>,
) -> Result<(), CanonError> {
    if let Some(layout) = layout {
        check_layout(ptr, layout.align, layout.size, memory.len())?;
    } else {
        check(ptr, t, memory.len())?;
    }
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
            let _ = memory.write_array(addr, &n.to_le_bytes());
        }
        ValType::I32 => {
            let _ = memory.store_i32(addr, v.as_i32());
        }
        ValType::U32 => {
            let _ = memory.store_i32(addr, unsigned_u32(v)? as i32);
        }
        ValType::I64 => {
            let _ = memory.store_i64(addr, v.as_i64());
        }
        ValType::U64 => {
            let _ = memory.store_i64(addr, unsigned_u64(v)? as i64);
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
            let bytes = packed.to_le_bytes();
            memory.write_bytes(
                addr,
                &bytes[..crate::canon_layout::flags_bytes(labels.len()) as usize],
            );
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
            let text = string_text(v);
            let bytes = text.as_bytes();
            let dest = realloc(bytes.len() as u32, 1).ok_or(CanonError::OutOfMemory {
                size: bytes.len() as u32,
            })?;
            write_bytes(memory, dest, bytes)?;
            store_pair_header(memory, addr, dest, bytes.len() as u32);
        }
        // `store_list` — same (ptr, length) shape, but length counts ELEMENTS
        // and each is stored at its own canonical stride.
        ValType::List(elem) => {
            if matches!(elem.as_ref(), ValType::U8) {
                let bytes = byte_items(v)?;
                let dest = if bytes.is_empty() {
                    0
                } else {
                    realloc(bytes.len() as u32, 1).ok_or(CanonError::OutOfMemory {
                        size: bytes.len() as u32,
                    })?
                };
                write_bytes(memory, dest, &bytes)?;
                store_pair_header(memory, addr, dest, bytes.len() as u32);
                return Ok(());
            }
            let items = array_items(v);
            let child = layout.and_then(|layout| match &layout.shape {
                crate::canon_layout::MemShape::Elements(child) => Some(child.as_ref()),
                _ => None,
            });
            let stride = child.map_or_else(|| elem_size(elem), |child| child.size);
            if stride == elem_size(elem) {
                if let Some(result) = store_scalar_list(memory, realloc, elem, &items, false) {
                    let (dest, count) = result?;
                    store_pair_header(memory, addr, dest, count);
                    return Ok(());
                }
            }
            let dest = if items.is_empty() {
                0
            } else {
                let bytes = stride.saturating_mul(items.len() as u32);
                realloc(
                    bytes,
                    child.map_or_else(|| alignment(elem), |child| child.align),
                )
                .ok_or(CanonError::OutOfMemory { size: bytes })?
            };
            for (i, item) in items.iter().enumerate() {
                store_impl(memory, realloc, item, elem, dest + stride * i as u32, child)?;
            }
            store_pair_header(memory, addr, dest, items.len() as u32);
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
            let child = layout.and_then(|layout| match &layout.shape {
                crate::canon_layout::MemShape::Elements(child) => Some(child.as_ref()),
                _ => None,
            });
            let stride = child.map_or_else(|| elem_size(elem), |child| child.size);
            for (i, item) in items.iter().enumerate() {
                store_impl(memory, realloc, item, elem, ptr + stride * i as u32, child)?;
            }
        }
        // `store_record` — each field aligned before it is placed, in
        // declaration order. Field lookup is by NAME, so a host value carrying
        // extra properties stores correctly and a missing one stores a default
        // rather than shifting every later field.
        ValType::Record(fields) => {
            let mut offset = 0u32;
            for (index, (name, field_ty)) in fields.iter().enumerate() {
                let child = layout.and_then(|layout| match &layout.shape {
                    crate::canon_layout::MemShape::Record(fields) => fields.get(index),
                    _ => None,
                });
                offset = child.map_or_else(|| align_to(offset, alignment(field_ty)), |(at, _)| *at);
                let field = record_field(v, name);
                store_impl(
                    memory,
                    realloc,
                    &field,
                    field_ty,
                    ptr + offset,
                    child.map(|(_, layout)| layout),
                )?;
                offset += child.map_or_else(|| elem_size(field_ty), |(_, layout)| layout.size);
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
                VariantCases::Pair([None, Some(inner.as_ref())]),
                layout,
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
                VariantCases::Pair([ok.as_deref(), err.as_deref()]),
                layout,
            )?;
        }
        ValType::Variant(cases) => {
            let (case, payload) = variant_case(v, cases)?;
            store_variant_parts(memory, realloc, ptr, case, &payload, VariantCases::Named(cases), layout)?;
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
    read_utf8_as(memory, ptr, len)
}

fn read_utf8_as<T>(
    memory: &crate::shared_memory::SharedMemory,
    ptr: usize,
    len: usize,
) -> Result<T, CanonError>
where
    T: for<'a> From<&'a str>,
{
    memory
        .map_read_range(ptr, len, |bytes| {
            std::str::from_utf8(bytes)
                .map(T::from)
                .map_err(|_| CanonError::Unsupported("string is not valid UTF-8"))
        })
        .map_err(|memory_len| CanonError::OutOfBounds {
            ptr: ptr as u32,
            size: len as u32,
            memory_len,
        })?
}

/// The discriminant and payload of a stored `variant` — the read counterpart of
/// [`store_variant_parts`], so the two share one view of where a payload lives.
enum VariantCases<'a> {
    Named(&'a [(String, Option<ValType>)]),
    Pair([Option<&'a ValType>; 2]),
}

impl VariantCases<'_> {
    fn len(&self) -> usize {
        match self {
            Self::Named(cases) => cases.len(),
            Self::Pair(_) => 2,
        }
    }

    fn payload_type(&self, case: usize) -> Option<&ValType> {
        match self {
            Self::Named(cases) => cases.get(case).and_then(|(_, ty)| ty.as_ref()),
            Self::Pair(cases) => cases.get(case).copied().flatten(),
        }
    }

    fn discriminant_size(&self) -> u32 {
        match self {
            Self::Named(cases) => crate::canon_layout::variant_discriminant_size(cases),
            Self::Pair(_) => 1,
        }
    }

    fn payload_offset(&self) -> u32 {
        match self {
            Self::Named(cases) => crate::canon_layout::variant_payload_offset(cases),
            Self::Pair(cases) => align_to(
                1,
                cases.iter().flatten().map(|ty| alignment(ty)).max().unwrap_or(1),
            ),
        }
    }
}

fn load_variant_parts(
    memory: &crate::shared_memory::SharedMemory,
    ptr: u32,
    cases: VariantCases<'_>,
    layout: Option<&crate::canon_layout::MemLayout>,
) -> Result<(u32, Value), CanonError> {
    let addr = ptr as usize;
    let variant = layout.and_then(|layout| match &layout.shape {
        crate::canon_layout::MemShape::Variant {
            tag_size,
            payload_offset,
            cases,
        } => Some((*tag_size, *payload_offset, cases.as_slice())),
        _ => None,
    });
    let case = match variant.map_or_else(
        || cases.discriminant_size(),
        |v| v.0,
    ) {
        1 => memory.load_u8(addr).unwrap_or(0) as u32,
        2 => load_u16(memory, addr) as u32,
        _ => memory.load_i32(addr).unwrap_or(0) as u32,
    };
    if case as usize >= cases.len() {
        return Err(CanonError::DiscriminantOutOfRange {
            got: case,
            cases: cases.len(),
        });
    }
    let payload = match cases.payload_type(case as usize) {
        Some(payload_ty) => {
            let offset = variant.map_or_else(
                || cases.payload_offset(),
                |v| v.1,
            );
            let child = variant
                .and_then(|v| v.2.get(case as usize))
                .and_then(Option::as_ref);
            load_impl(memory, payload_ty, ptr + offset, child)?
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
    memory.write_bytes(ptr as usize, bytes);
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

const MAX_LIST_BYTE_LENGTH: usize = (1 << 28) - 1;

fn check_byte_list_len(bytes: usize) -> Result<(), CanonError> {
    if bytes > MAX_LIST_BYTE_LENGTH {
        return Err(CanonError::ListByteLengthExceeded { bytes });
    }
    Ok(())
}

fn byte_items(v: &Value) -> Result<Vec<u8>, CanonError> {
    use crate::value::ObjectKind;
    if let Value::Object(object) = v {
        let object = object.lock().unwrap();
        if let ObjectKind::Array(items) = &object.kind {
            check_byte_list_len(items.len())?;
            return Ok(items.iter().map(|item| item.as_i32() as u8).collect());
        }
    }
    Ok(Vec::new())
}

/// Decode scalar lists under one checked guard; only owned values escape it.
fn load_scalar_items(
    memory: &crate::shared_memory::SharedMemory,
    elem: &ValType,
    ptr: u32,
    count: u32,
) -> Option<Result<Value, CanonError>> {
    let stride = match elem {
        ValType::Bool | ValType::S8 | ValType::U8 => 1usize,
        ValType::S16 | ValType::U16 => 2,
        ValType::I32 | ValType::U32 | ValType::F32 => 4,
        ValType::I64 | ValType::U64 | ValType::F64 => 8,
        _ => return None,
    };
    Some((|| {
        let bytes = (count as usize).checked_mul(stride).unwrap_or(usize::MAX);
        check_byte_list_len(bytes)?;
        if ptr as usize % stride != 0 {
            return Err(CanonError::Misaligned { ptr, required: stride as u32 });
        }
        let items = memory.map_read_range(ptr as usize, bytes, |range| {
            range.chunks_exact(stride).map(|item| match elem {
                ValType::Bool => Value::Bool(item[0] != 0),
                ValType::S8 => Value::I32(item[0] as i8 as i32),
                ValType::U8 => Value::I32(item[0] as i32),
                ValType::S16 => Value::I32(i16::from_le_bytes(item.try_into().unwrap()) as i32),
                ValType::U16 => Value::I32(u16::from_le_bytes(item.try_into().unwrap()) as i32),
                ValType::I32 => Value::I32(i32::from_le_bytes(item.try_into().unwrap())),
                ValType::U32 => Value::F64(u32::from_le_bytes(item.try_into().unwrap()) as f64),
                ValType::I64 => Value::I64(i64::from_le_bytes(item.try_into().unwrap())),
                ValType::U64 => Value::bigint_u64(u64::from_le_bytes(item.try_into().unwrap())),
                ValType::F32 => Value::F64(f32::from_bits(u32::from_le_bytes(item.try_into().unwrap())) as f64),
                ValType::F64 => Value::F64(f64::from_bits(u64::from_le_bytes(item.try_into().unwrap()))),
                _ => unreachable!("non-scalar list entered scalar decoder"),
            }).collect::<Vec<_>>()
        }).map_err(|memory_len| CanonError::OutOfBounds { ptr, size: bytes as u32, memory_len })?;
        Ok(Value::Object(crate::heap::alloc(crate::value::Object::new_array(items))))
    })())
}

/// Decode a list of flat numeric records from one memory snapshot. Invalid
/// ranges and non-scalar fields use the recursive path so its trap ordering is
/// unchanged.
fn load_plain_record_items(
    memory: &crate::shared_memory::SharedMemory,
    fields: &[(String, ValType)],
    layout: &crate::canon_layout::MemLayout,
    ptr: u32,
    count: u32,
) -> Option<Value> {
    let crate::canon_layout::MemShape::Record(parts) = &layout.shape else {
        return None;
    };
    if layout.size == 0 || parts.len() != fields.len() {
        return None;
    }
    for ((_, ty), (offset, field_layout)) in fields.iter().zip(parts) {
        if !matches!(
            ty,
            ValType::Bool
                | ValType::S8
                | ValType::U8
                | ValType::S16
                | ValType::U16
                | ValType::I32
                | ValType::U32
                | ValType::I64
                | ValType::U64
                | ValType::F32
                | ValType::F64
        ) || offset.checked_add(field_layout.size)? > layout.size
        {
            return None;
        }
    }
    let start = ptr as usize;
    let size = layout.size as usize;
    let byte_len = (count as usize).checked_mul(size)?;
    if start % layout.align as usize != 0
        || start.checked_add(byte_len)? > memory.len()
    {
        return None;
    }
    let snapshot = memory
        .map_read_range(start, byte_len, |bytes| bytes.to_vec())
        .ok()?;
    Some(decode_plain_record_items(fields, parts, size, &snapshot))
}

fn decode_plain_record_items(
    fields: &[(String, ValType)],
    parts: &[(u32, crate::canon_layout::MemLayout)],
    size: usize,
    bytes: &[u8],
) -> Value {
    decode_plain_record_items_with_map::<true, true>(fields, parts, size, bytes)
}

#[cfg(test)]
fn decode_plain_record_items_with_capacity<const RESERVE: bool>(
    fields: &[(String, ValType)],
    parts: &[(u32, crate::canon_layout::MemLayout)],
    size: usize,
    bytes: &[u8],
) -> Value {
    decode_plain_record_items_with_map::<RESERVE, false>(fields, parts, size, bytes)
}

fn decode_plain_record_items_with_map<const RESERVE: bool, const TEMPLATE: bool>(
    fields: &[(String, ValType)],
    parts: &[(u32, crate::canon_layout::MemLayout)],
    size: usize,
    bytes: &[u8],
) -> Value {
    if TEMPLATE && bytes.len() / size <= 1 {
        return decode_plain_record_items_with_map::<RESERVE, false>(fields, parts, size, bytes);
    }
    let mut template = indexmap::IndexMap::new();
    if TEMPLATE {
        template.reserve(fields.len());
        for (name, _) in fields {
            template.insert(name.clone(), Value::Undefined);
        }
        // Valid canonical records have unique names; retain the prior behavior
        // for direct internal callers supplying an unvalidated record type.
        if template.len() != fields.len() {
            return decode_plain_record_items_with_map::<RESERVE, false>(fields, parts, size, bytes);
        }
    }
    let items = bytes
        .chunks_exact(size)
        .map(|record| {
            let mut object = crate::value::Object::new();
            if TEMPLATE {
                object.properties = template.clone();
            } else if RESERVE && fields.len() > 2 {
                object.properties.reserve(fields.len());
            }
            for (index, ((name, ty), (offset, _))) in fields.iter().zip(parts).enumerate() {
                let at = *offset as usize;
                let value = match ty {
                    ValType::Bool => Value::Bool(record[at] != 0),
                    ValType::S8 => Value::I32(record[at] as i8 as i32),
                    ValType::U8 => Value::I32(record[at] as i32),
                    ValType::S16 => Value::I32(i16::from_le_bytes(record[at..at + 2].try_into().unwrap()) as i32),
                    ValType::U16 => Value::I32(u16::from_le_bytes(record[at..at + 2].try_into().unwrap()) as i32),
                    ValType::I32 => Value::I32(i32::from_le_bytes(record[at..at + 4].try_into().unwrap())),
                    ValType::U32 => Value::F64(u32::from_le_bytes(record[at..at + 4].try_into().unwrap()) as f64),
                    ValType::I64 => Value::I64(i64::from_le_bytes(record[at..at + 8].try_into().unwrap())),
                    ValType::U64 => Value::bigint_u64(u64::from_le_bytes(record[at..at + 8].try_into().unwrap())),
                    ValType::F32 => Value::F64(f32::from_bits(u32::from_le_bytes(record[at..at + 4].try_into().unwrap())) as f64),
                    ValType::F64 => Value::F64(f64::from_bits(u64::from_le_bytes(record[at..at + 8].try_into().unwrap()))),
                    _ => unreachable!("record field shape was checked"),
                };
                if TEMPLATE {
                    *object.properties.get_index_mut(index).unwrap().1 = value;
                } else {
                    object.properties.insert(name.as_str().into(), value);
                }
            }
            Value::Object(crate::heap::alloc(object))
        })
        .collect();
    Value::Object(crate::heap::alloc(
        crate::value::Object::new_array(items),
    ))
}

/// Keep the source snapshot outside guest realloc and the destination guard.
/// Object/string coercions stay on the recursive path, avoiding lock reentry.
fn store_scalar_list(
    memory: &crate::shared_memory::SharedMemory,
    realloc: Realloc<'_>,
    elem: &ValType,
    items: &[Value],
    allocate_empty: bool,
) -> Option<Result<(u32, u32), CanonError>> {
    let stride = match elem {
        ValType::Bool | ValType::S8 => 1usize,
        ValType::S16 | ValType::U16 => 2,
        ValType::I32 | ValType::U32 | ValType::F32 => 4,
        ValType::I64 | ValType::U64 | ValType::F64 => 8,
        _ => return None,
    };
    if matches!(elem, ValType::U32 | ValType::U64) {
        for value in items {
            let result = if matches!(elem, ValType::U32) {
                unsigned_u32(value).map(|_| ())
            } else {
                unsigned_u64(value).map(|_| ())
            };
            if let Err(error) = result {
                return Some(Err(error));
            }
        }
    } else if !items.iter().all(|v| matches!(v,
        Value::I32(_) | Value::I64(_) | Value::F64(_) | Value::Bool(_) | Value::Null | Value::Undefined
    )) {
        return None;
    }
    Some((|| {
        let bytes = items.len().checked_mul(stride).unwrap_or(usize::MAX);
        if bytes > u32::MAX as usize {
            return Err(CanonError::ListByteLengthExceeded { bytes });
        }
        let size = bytes as u32;
        let dest = if items.is_empty() && !allocate_empty {
            0
        } else {
            realloc(size.max(u32::from(allocate_empty)), stride as u32)
                .ok_or(CanonError::OutOfMemory { size })?
        };
        if dest as usize % stride != 0 {
            return Err(CanonError::Misaligned { ptr: dest, required: stride as u32 });
        }
        memory.map_write_range(dest as usize, bytes, |range| {
            for (out, value) in range.chunks_exact_mut(stride).zip(items) {
                match elem {
                    ValType::Bool => out[0] = u8::from(value.as_bool()),
                    ValType::S8 => out[0] = value.as_i32() as u8,
                    ValType::S16 | ValType::U16 => out.copy_from_slice(&(value.as_i32() as u16).to_le_bytes()),
                    ValType::I32 => out.copy_from_slice(&value.as_i32().to_le_bytes()),
                    ValType::U32 => out.copy_from_slice(&unsigned_u32(value).expect("validated unsigned list item").to_le_bytes()),
                    ValType::I64 => out.copy_from_slice(&value.as_i64().to_le_bytes()),
                    ValType::U64 => out.copy_from_slice(&unsigned_u64(value).expect("validated unsigned list item").to_le_bytes()),
                    ValType::F32 => out.copy_from_slice(&(value.as_f64() as f32).to_bits().to_le_bytes()),
                    ValType::F64 => out.copy_from_slice(&value.as_f64().to_bits().to_le_bytes()),
                    _ => unreachable!("non-scalar list entered scalar encoder"),
                }
            }
        }).map_err(|memory_len| CanonError::OutOfBounds { ptr: dest, size, memory_len })?;
        Ok((dest, items.len() as u32))
    })())
}

fn load_byte_items(
    memory: &crate::shared_memory::SharedMemory,
    ptr: u32,
    len: u32,
) -> Result<Value, CanonError> {
    check_byte_list_len(len as usize)?;
    let items = memory
        .map_read_range(ptr as usize, len as usize, |bytes| {
            bytes.iter().map(|byte| Value::I32(*byte as i32)).collect::<Vec<_>>()
        })
        .map_err(|memory_len| CanonError::OutOfBounds {
            ptr,
            size: len,
            memory_len,
        })?;
    Ok(Value::Object(crate::heap::alloc(
        crate::value::Object::new_array(items),
    )))
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
    cases: VariantCases<'_>,
    layout: Option<&crate::canon_layout::MemLayout>,
) -> Result<(), CanonError> {
    let addr = ptr as usize;
    let variant = layout.and_then(|layout| match &layout.shape {
        crate::canon_layout::MemShape::Variant {
            tag_size,
            payload_offset,
            cases,
        } => Some((*tag_size, *payload_offset, cases.as_slice())),
        _ => None,
    });
    match variant.map_or_else(
        || cases.discriminant_size(),
        |v| v.0,
    ) {
        1 => {
            let _ = memory.store_u8(addr, case as u8);
        }
        2 => {
            let _ = memory.write_array(addr, &(case as u16).to_le_bytes());
        }
        _ => {
            let _ = memory.store_i32(addr, case as i32);
        }
    }
    if let Some(payload_ty) = cases.payload_type(case as usize) {
        let offset = variant.map_or_else(
            || cases.payload_offset(),
            |v| v.1,
        );
        let child = variant
            .and_then(|v| v.2.get(case as usize))
            .and_then(Option::as_ref);
        store_impl(memory, realloc, payload, payload_ty, ptr + offset, child)?;
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
    Err(CanonError::Unsupported(
        "variant case not named by the value",
    ))
}

/// `load(cx, ptr, t)` — read one value of type `t` from `ptr`.
pub fn load(
    memory: &crate::shared_memory::SharedMemory,
    t: &ValType,
    ptr: u32,
) -> Result<Value, CanonError> {
    load_impl(memory, t, ptr, None)
}

pub(crate) fn load_with_layout(
    memory: &crate::shared_memory::SharedMemory,
    t: &ValType,
    ptr: u32,
    layout: &crate::canon_layout::MemLayout,
) -> Result<Value, CanonError> {
    load_impl(memory, t, ptr, Some(layout))
}

fn load_impl(
    memory: &crate::shared_memory::SharedMemory,
    t: &ValType,
    ptr: u32,
    layout: Option<&crate::canon_layout::MemLayout>,
) -> Result<Value, CanonError> {
    if let Some(layout) = layout {
        check_layout(ptr, layout.align, layout.size, memory.len())?;
    } else {
        check(ptr, t, memory.len())?;
    }
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
        ValType::U32 => Value::F64(memory.load_i32(addr).unwrap_or(0) as u32 as f64),
        ValType::I64 => Value::I64(memory.load_i64(addr).unwrap_or(0)),
        ValType::U64 => Value::bigint_u64(memory.load_i64(addr).unwrap_or(0) as u64),
        ValType::F32 => {
            Value::F64(f32::from_bits(memory.load_i32(addr).unwrap_or(0) as u32) as f64)
        }
        ValType::F64 => Value::F64(memory.load_f64(addr).unwrap_or(0.0)),
        ValType::Char => {
            let raw = memory.load_i32(addr).unwrap_or(0) as u32;
            Value::String(std::sync::Arc::from(
                scalar_to_char(raw)?.to_string().as_str(),
            ))
        }
        // `load_flags` = `unpack_flags_from_int(load_int(ptr, elem_size))` —
        // read only the packed width, then one bit per label.
        ValType::Flags(labels) => {
            let mut packed: u32 = 0;
            let count = crate::canon_layout::flags_bytes(labels.len()) as usize;
            let mut bytes = [0u8; 4];
            memory.read_bytes(addr, &mut bytes[..count]);
            for (k, byte) in bytes[..count].iter().enumerate() {
                packed |= (*byte as u32) << (8 * k);
            }
            unpack_flags(packed, labels)
        }
        ValType::Own(_)
        | ValType::Borrow(_)
        | ValType::Stream(_)
        | ValType::Future(_)
        | ValType::ErrorContext => Value::I32(memory.load_i32(addr).unwrap_or(0)),
        // `load_string` — (ptr, length) where length is BYTES of UTF-8.
        ValType::String => {
            let (at, len) = load_pair_header(memory, addr);
            Value::String(read_utf8_as(memory, at as usize, len as usize)?)
        }
        // `load_list` — (ptr, length) where length counts ELEMENTS.
        ValType::List(elem) => {
            let (at, count) = load_pair_header(memory, addr);
            if matches!(elem.as_ref(), ValType::U8) {
                return load_byte_items(memory, at, count);
            }
            let child = layout.and_then(|layout| match &layout.shape {
                crate::canon_layout::MemShape::Elements(child) => Some(child.as_ref()),
                _ => None,
            });
            let stride = child.map_or_else(|| elem_size(elem), |child| child.size);
            if stride == elem_size(elem) {
                if let Some(result) = load_scalar_items(memory, elem, at, count) {
                    return result;
                }
            }
            if let (ValType::Record(fields), Some(child)) = (elem.as_ref(), child) {
                if let Some(result) = load_plain_record_items(memory, fields, child, at, count) {
                    return Ok(result);
                }
            }
            load_list_items_recursive(memory, elem, at, count, stride, child)?
        }
        // 🔧 `load_list_from_valid_range(cx, ptr, N, t)` — N elements read
        // in place. The count comes from the TYPE, not from memory.
        ValType::ListFixed(elem, n) => {
            let child = layout.and_then(|layout| match &layout.shape {
                crate::canon_layout::MemShape::Elements(child) => Some(child.as_ref()),
                _ => None,
            });
            let stride = child.map_or_else(|| elem_size(elem), |child| child.size);
            let mut items = Vec::with_capacity(*n as usize);
            for i in 0..*n {
                items.push(load_impl(memory, elem, ptr + stride * i, child)?);
            }
            Value::Object(crate::heap::alloc(crate::value::Object::new_array(items)))
        }
        // `load_record` — fields in declaration order, each at its aligned offset.
        ValType::Record(fields) => {
            let mut object = crate::value::Object::new();
            let mut offset = 0u32;
            for (index, (name, field_ty)) in fields.iter().enumerate() {
                let child = layout.and_then(|layout| match &layout.shape {
                    crate::canon_layout::MemShape::Record(fields) => fields.get(index),
                    _ => None,
                });
                offset = child.map_or_else(|| align_to(offset, alignment(field_ty)), |(at, _)| *at);
                let value = load_impl(
                    memory,
                    field_ty,
                    ptr + offset,
                    child.map(|(_, layout)| layout),
                )?;
                object.properties.insert(name.as_str().into(), value);
                offset += child.map_or_else(|| elem_size(field_ty), |(_, layout)| layout.size);
            }
            Value::Object(crate::heap::alloc(object))
        }
        // `option` and `result` despecialise to `variant`, so they load through
        // the same discriminant arithmetic rather than repeating it.
        ValType::Option(inner) => {
            let (case, payload) = load_variant_parts(
                memory,
                ptr,
                VariantCases::Pair([None, Some(inner.as_ref())]),
                layout,
            )?;
            // `none` is `Value::Null`, which is how every host here spells it.
            if case == 0 { Value::Null } else { payload }
        }
        ValType::Result(ok, err) => {
            let (case, payload) = load_variant_parts(
                memory,
                ptr,
                VariantCases::Pair([ok.as_deref(), err.as_deref()]),
                layout,
            )?;
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
            let (case, payload) = load_variant_parts(memory, ptr, VariantCases::Named(cases), layout)?;
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
        ValType::Any => return Err(CanonError::Unsupported("any (not a component type)")),
    })
}

fn load_list_items_recursive(
    memory: &crate::shared_memory::SharedMemory,
    elem: &ValType,
    at: u32,
    count: u32,
    stride: u32,
    child: Option<&crate::canon_layout::MemLayout>,
) -> Result<Value, CanonError> {
    let mut items = Vec::with_capacity(count as usize);
    for i in 0..count {
        items.push(load_impl(memory, elem, at + stride * i, child)?);
    }
    Ok(Value::Object(crate::heap::alloc(
        crate::value::Object::new_array(items),
    )))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::shared_memory::SharedMemory;

    #[test]
    fn unsigned_scalars_preserve_full_range_in_memory() {
        let memory = SharedMemory::new(16);
        for bits in [0, 1 << 31, u32::MAX] {
            store(&memory, &Value::F64(bits as f64), &ValType::U32, 0).unwrap();
            assert_eq!(memory.load_i32(0).unwrap() as u32, bits);
            assert!(matches!(load(&memory, &ValType::U32, 0).unwrap(), Value::F64(n) if n == bits as f64));
        }
        for bits in [0, 1 << 63, u64::MAX] {
            store(&memory, &Value::bigint_u64(bits), &ValType::U64, 8).unwrap();
            assert_eq!(memory.load_i64(8).unwrap() as u64, bits);
            assert!(matches!(load(&memory, &ValType::U64, 8).unwrap(), Value::BigInt(n) if n.to_u64_wrapping() == bits && !n.is_negative()));
        }
        assert!(matches!(store(&memory, &Value::F64(4_294_967_296.0), &ValType::U32, 0), Err(CanonError::UnsignedOutOfRange { bits: 32 })));
        assert!(matches!(store(&memory, &Value::F64(1.5), &ValType::U32, 0), Err(CanonError::UnsignedOutOfRange { bits: 32 })));
        assert!(matches!(store(&memory, &Value::F64(-1.0), &ValType::U32, 0), Err(CanonError::UnsignedOutOfRange { bits: 32 })));
        assert!(matches!(store(&memory, &Value::I32(-1), &ValType::U32, 0), Err(CanonError::UnsignedOutOfRange { bits: 32 })));
        assert!(matches!(store(&memory, &Value::F64(18_446_744_073_709_551_616.0), &ValType::U64, 8), Err(CanonError::UnsignedOutOfRange { bits: 64 })));
        let beyond_u64 = Value::bigint(crate::bigint::BigIntVal::parse("18446744073709551616").unwrap());
        assert!(matches!(store(&memory, &beyond_u64, &ValType::U64, 8), Err(CanonError::UnsignedOutOfRange { bits: 64 })));
        assert!(matches!(store(&memory, &Value::I64(-1), &ValType::U64, 8), Err(CanonError::UnsignedOutOfRange { bits: 64 })));
        assert_eq!(memory.load_i32(0).unwrap() as u32, u32::MAX);
        assert_eq!(memory.load_i64(8).unwrap() as u64, u64::MAX);
    }

    #[test]
    fn pair_headers_preserve_unsigned_words_and_little_endian_layout() {
        let memory = SharedMemory::new(16);
        for (ptr, len) in [(0, 0), (0x81234567, 0x89abcdef), (u32::MAX, u32::MAX)] {
            store_pair_header(&memory, 4, ptr, len);
            assert_eq!(load_pair_header(&memory, 4), (ptr, len));
            assert_eq!(memory.load_i32(4).unwrap() as u32, ptr);
            assert_eq!(memory.load_i32(8).unwrap() as u32, len);
        }
    }

    #[test]
    fn pair_header_checks_precede_allocator_and_preserve_memory() {
        let memory = SharedMemory::new(16);
        memory.write_array(0, &[0xa5u8; 16]).unwrap();
        let value = Value::String(std::sync::Arc::from("text"));
        for ptr in [1u32, 12u32] {
            let mut called = false;
            let mut allocator = |_size, _align| { called = true; Some(8) };
            assert!(store_with(&memory, &mut allocator, &value, &ValType::String, ptr).is_err());
            assert!(!called);
            assert_eq!(memory.read_array::<16>(0).unwrap(), [0xa5; 16]);
        }
    }

    #[test]
    #[ignore = "alternating loaded-machine pointer-length header measurement"]
    fn pair_header_loaded_benchmark() {
        let memory = SharedMemory::new(16);
        store_pair_header(&memory, 4, 0x81234567, 0x89abcdef);
        for read in [true, false] {
            let run = |current| {
                let start = std::time::Instant::now();
                for _ in 0..1_000_000 {
                    let memory = std::hint::black_box(&memory);
                    if read {
                        let pair = if current { load_pair_header(memory, 4) }
                            else { (memory.load_i32(4).unwrap() as u32, memory.load_i32(8).unwrap() as u32) };
                        std::hint::black_box(pair);
                    } else if current {
                        store_pair_header(memory, 4, 0x81234567, 0x89abcdef);
                    } else {
                        memory.store_i32(4, 0x81234567u32 as i32).unwrap();
                        memory.store_i32(8, 0x89abcdefu32 as i32).unwrap();
                    }
                }
                start.elapsed().as_nanos()
            };
            run(false); run(true);
            let mut before = Vec::with_capacity(21);
            let mut after = Vec::with_capacity(21);
            let mut ratios = Vec::with_capacity(21);
            for round in 0..21 {
                let (a, b) = if round % 2 == 0 { (run(false), run(true)) }
                    else { let b = run(true); (run(false), b) };
                before.push(a); after.push(b); ratios.push(a as f64 / b as f64);
            }
            before.sort_unstable(); after.sort_unstable(); ratios.sort_by(f64::total_cmp);
            assert_eq!(load_pair_header(&memory, 4), (0x81234567, 0x89abcdef));
            eprintln!("pair-header read={read} before_p50_ns={} after_p50_ns={} before_p95_ns={} after_p95_ns={} ratio_p10={:.3} ratio_p50={:.3} ratio_p90={:.3}", before[10], after[10], before[19], after[19], ratios[2], ratios[10], ratios[18]);
        }
    }

    #[test]
    fn utf8_range_results_own_their_bytes() {
        let (text, shared) = {
            let memory = SharedMemory::new(64);
            memory.write_array(8, &[0x61u8, 0, 0xc3, 0xa9, 0xf0, 0x9f, 0x99, 0x82]).unwrap();
            let text = read_utf8(&memory, 8, 8).unwrap();
            let shared: std::sync::Arc<str> = read_utf8_as(&memory, 8, 8).unwrap();
            memory.write_array(8, &[0u8; 8]).unwrap();
            (text, shared)
        };
        assert_eq!(text, "a\0\u{e9}\u{1f642}");
        assert_eq!(shared.as_ref(), text);
    }

    #[test]
    fn utf8_range_preserves_invalid_encoding_and_empty_bounds_errors() {
        let memory = SharedMemory::new(16);
        memory.write_array(0, &[0xffu8, 0xff]).unwrap();
        assert!(matches!(read_utf8(&memory, 0, 2), Err(CanonError::Unsupported(_))));
        assert_eq!(read_utf8(&memory, 16, 0).unwrap(), "");
        for (ptr, len) in [(17usize, 0usize), (usize::MAX, 1)] {
            assert!(matches!(read_utf8(&memory, ptr, len), Err(CanonError::OutOfBounds { .. })));
        }
    }

    #[test]
    fn byte_list_lift_is_independent_of_later_memory_writes() {
        let memory = SharedMemory::new(64);
        memory.write_array(32, &[0u8, 127, 128, 255]).unwrap();
        let ty = ValType::List(Box::new(ValType::U8));
        let value = load_pair_public(&memory, &ty, 32, 4).unwrap();
        memory.write_array(32, &[9u8; 4]).unwrap();
        assert_eq!(array_items(&value), vec![Value::I32(0), Value::I32(127), Value::I32(128), Value::I32(255)]);
    }

    #[test]
    fn string_lowering_borrows_text_and_preserves_display_fallback() {
        let value = Value::String(std::sync::Arc::from("a\0\u{e9}\u{1f642}"));
        assert!(matches!(string_text(&value), std::borrow::Cow::Borrowed(_)));
        assert_eq!(string_text(&value).as_bytes(), value.as_str().as_bytes());
        assert_eq!(string_text(&Value::I32(-42)), "-42");
        assert_eq!(
            string_text(&Value::Symbol(std::sync::Arc::from("label"))),
            "Symbol(label)"
        );
    }

    #[test]
    fn string_lowering_keeps_utf8_lengths_and_empty_allocator_requests() {
        for text in ["", "a\0\u{e9}\u{1f642}"] {
            let memory = SharedMemory::new(128);
            let value = Value::String(std::sync::Arc::from(text));
            let mut requested = Vec::new();
            let mut allocator = |size, align| {
                requested.push((size, align));
                Some(64)
            };
            store_with(&memory, &mut allocator, &value, &ValType::String, 0).unwrap();
            assert_eq!(requested, vec![(text.len() as u32, 1)]);
            assert_eq!(memory.load_i32(0).unwrap(), 64);
            assert_eq!(memory.load_i32(4).unwrap(), text.len() as i32);
            assert_eq!(load(&memory, &ValType::String, 0).unwrap().as_str(), text);

            let mut requested = Vec::new();
            let mut allocator_fn = |size, align| {
                requested.push((size, align));
                Some(64)
            };
            let mut allocator: Realloc<'_> = &mut allocator_fn;
            let pair = store_pair_public(&memory, &mut allocator, &value, &ValType::String)
                .unwrap();
            assert_eq!(pair, (64, text.len() as u32));
            assert_eq!(requested, vec![(text.len().max(1) as u32, 1)]);
            assert_eq!(load_pair_public(&memory, &ValType::String, pair.0, pair.1).unwrap().as_str(), text);
        }
    }

    #[test]
    fn string_lowering_failure_does_not_overwrite_header() {
        let memory = SharedMemory::new(128);
        memory.write_array(0, &[0xa5; 8]).unwrap();
        let value = Value::String(std::sync::Arc::from("text"));
        let mut allocator = |_size, _align| None;
        assert!(matches!(
            store_with(&memory, &mut allocator, &value, &ValType::String, 0),
            Err(CanonError::OutOfMemory { size: 4 })
        ));
        assert_eq!(memory.read_array::<8>(0).unwrap(), [0xa5; 8]);
        let mut allocator = |_size, _align| Some(127);
        assert!(matches!(
            store_with(&memory, &mut allocator, &value, &ValType::String, 0),
            Err(CanonError::OutOfBounds { .. })
        ));
        assert_eq!(memory.read_array::<8>(0).unwrap(), [0xa5; 8]);
    }

    #[test]
    #[ignore = "alternating loaded-machine canonical string lowering measurement"]
    fn string_lowering_loaded_benchmark() {
        fn previous(memory: &SharedMemory, realloc: Realloc<'_>, v: &Value) {
            check(0, &ValType::String, memory.len()).unwrap();
            let text = format!("{v}");
            let bytes = text.as_bytes();
            let dest = realloc(bytes.len() as u32, 1).unwrap();
            write_bytes(memory, dest, bytes).unwrap();
            memory.store_i32(0, dest as i32).unwrap();
            memory.store_i32(PTR_SIZE as usize, bytes.len() as i32).unwrap();
        }
        for (size, count) in [(4096usize, 20_000usize), (65536, 2000)] {
            let memory = SharedMemory::new(size + 64);
            let value = Value::String(std::sync::Arc::from("x".repeat(size)));
            let run = |current| {
                let mut allocator = |_size, _align| Some(32);
                let start = std::time::Instant::now();
                for _ in 0..count {
                    if current {
                        store_with(&memory, &mut allocator, std::hint::black_box(&value), &ValType::String, 0).unwrap();
                    } else {
                        previous(&memory, &mut allocator, std::hint::black_box(&value));
                    }
                }
                start.elapsed().as_nanos()
            };
            run(false);
            run(true);
            let mut previous = Vec::with_capacity(21);
            let mut current = Vec::with_capacity(21);
            let mut ratios = Vec::with_capacity(21);
            for round in 0..21 {
                let (before, after) = if round % 2 == 0 { (run(false), run(true)) }
                    else { let after = run(true); (run(false), after) };
                previous.push(before);
                current.push(after);
                ratios.push(before as f64 / after as f64);
            }
            previous.sort_unstable();
            current.sort_unstable();
            ratios.sort_by(f64::total_cmp);
            assert_eq!(memory.load_i32(0).unwrap(), 32);
            assert_eq!(memory.load_i32(4).unwrap(), size as i32);
            assert_eq!(read_utf8(&memory, 32, size).unwrap(), value.as_str());
            eprintln!("canonical string size={size} count={count}: previous_p50_ns={} current_p50_ns={} previous_p95_ns={} current_p95_ns={} paired_ratio_p10={:.3} paired_ratio_p50={:.3} paired_ratio_p90={:.3}", previous[10], current[10], previous[19], current[19], ratios[2], ratios[10], ratios[18]);
        }
    }

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
    fn canonical_variant_tags_reject_out_of_range_values() {
        let memory = mem();
        let types = [
            ValType::Option(Box::new(ValType::I32)),
            ValType::Result(Some(Box::new(ValType::I32)), Some(Box::new(ValType::I32))),
            ValType::Variant(vec![("first".into(), None), ("second".into(), None)]),
        ];
        for ty in &types {
            let layout = crate::canon_layout::MemLayout::new(ty);
            for tag in [2, u8::MAX] {
                memory.store_u8(0, tag).unwrap();
                for result in [load(&memory, ty, 0), load_with_layout(&memory, ty, 0, &layout)] {
                    assert!(matches!(result, Err(CanonError::DiscriminantOutOfRange { got, cases: 2 }) if got == tag as u32));
                }
            }
            for tag in [0, 1] {
                memory.store_u8(0, tag).unwrap();
                assert!(load(&memory, ty, 0).is_ok());
                assert!(load_with_layout(&memory, ty, 0, &layout).is_ok());
            }
        }
        memory.store_u8(0, 1).unwrap();
        memory.store_i32(4, 42).unwrap();
        let option = &types[0];
        assert_eq!(load(&memory, option, 0).unwrap().as_i32(), 42);
        assert_eq!(
            load_with_layout(&memory, option, 0, &crate::canon_layout::MemLayout::new(option))
                .unwrap()
                .as_i32(),
            42
        );

        let wide = ValType::Variant((0..257).map(|i| (format!("case{i}"), None)).collect());
        let layout = crate::canon_layout::MemLayout::new(&wide);
        memory.store_u8(0, 0).unwrap();
        memory.store_u8(1, 1).unwrap();
        assert!(load(&memory, &wide, 0).is_ok());
        assert!(load_with_layout(&memory, &wide, 0, &layout).is_ok());
        memory.store_u8(0, 1).unwrap();
        for result in [load(&memory, &wide, 0), load_with_layout(&memory, &wide, 0, &layout)] {
            assert!(matches!(result, Err(CanonError::DiscriminantOutOfRange { got: 257, cases: 257 })));
        }
    }

    #[test]
    fn borrowed_two_case_types_keep_canonical_payload_offsets() {
        let memory = mem();
        let option = ValType::Option(Box::new(ValType::I64));
        let result = ValType::Result(Some(Box::new(ValType::I64)), Some(Box::new(ValType::I64)));
        let mut error = crate::value::Object::new();
        error.properties.insert("__wasi_error".into(), Value::I64(73));
        let values = [
            (&option, Value::I64(42), 1u8, 42i64),
            (&result, Value::I64(42), 0, 42),
            (&result, Value::Object(crate::heap::alloc(error)), 1, 73),
        ];
        for (ty, value, tag, payload) in &values {
            let layout = crate::canon_layout::MemLayout::new(ty);
            for cached in [false, true] {
                memory.store_i64(0, 0).unwrap();
                let mut alloc = no_realloc;
                if cached {
                    store_with_layout(&memory, &mut alloc, value, ty, 0, &layout).unwrap();
                } else {
                    store_with(&memory, &mut alloc, value, ty, 0).unwrap();
                }
                assert_eq!(memory.load_u8(0).unwrap(), *tag);
                assert_eq!(memory.load_i64(8).unwrap(), *payload);
                let loaded = if cached {
                    load_with_layout(&memory, ty, 0, &layout).unwrap()
                } else {
                    load(&memory, ty, 0).unwrap()
                };
                if *tag == 1 && matches!(ty, ValType::Result(..)) {
                    let Value::Object(object) = loaded else { panic!("result error must be an object") };
                    assert_eq!(object.lock().unwrap().properties.get("__wasi_error").unwrap().as_i64(), *payload);
                } else {
                    assert_eq!(loaded.as_i64(), *payload);
                }
            }
        }
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

    #[test]
    fn cached_layout_matches_uncached_memory_and_traps() {
        let ty = ValType::Record(vec![
            (
                "items".into(),
                ValType::ListFixed(Box::new(ValType::U16), 3),
            ),
            (
                "choice".into(),
                ValType::Variant(vec![
                    ("empty".into(), None),
                    ("number".into(), Some(ValType::I64)),
                ]),
            ),
        ]);
        let layout = crate::canon_layout::MemLayout::new(&ty);
        assert_eq!((layout.size, layout.align), (24, 8));
        let mut choice = crate::value::Object::new();
        choice
            .properties
            .insert("tag".into(), Value::String("number".into()));
        choice.properties.insert("val".into(), Value::I64(93));
        let mut value = crate::value::Object::new();
        value.properties.insert(
            "items".into(),
            Value::Object(crate::heap::alloc(crate::value::Object::new_array(vec![
                Value::I32(2),
                Value::I32(65535),
                Value::I32(4),
            ]))),
        );
        value
            .properties
            .insert("choice".into(), Value::Object(crate::heap::alloc(choice)));
        let value = Value::Object(crate::heap::alloc(value));
        let uncached = SharedMemory::new(64);
        let cached = SharedMemory::new(64);
        let mut no_alloc = no_realloc;
        store_with(&uncached, &mut no_alloc, &value, &ty, 0).unwrap();
        store_with_layout(&cached, &mut no_alloc, &value, &ty, 0, &layout).unwrap();
        let mut a = [0; 64];
        let mut b = [0; 64];
        uncached.read_bytes(0, &mut a);
        cached.read_bytes(0, &mut b);
        assert_eq!(a, b);
        assert_eq!(
            format!("{:?}", load(&uncached, &ty, 0).unwrap()),
            format!("{:?}", load_with_layout(&cached, &ty, 0, &layout).unwrap())
        );

        for ptr in [1, 48] {
            let plain = load(&uncached, &ty, ptr).unwrap_err();
            let fast = load_with_layout(&uncached, &ty, ptr, &layout).unwrap_err();
            assert_eq!(format!("{plain:?}"), format!("{fast:?}"));
        }
        uncached.store_u8(8, 3).unwrap();
        cached.store_u8(8, 3).unwrap();
        let plain = load(&uncached, &ty, 0).unwrap_err();
        let fast = load_with_layout(&cached, &ty, 0, &layout).unwrap_err();
        assert_eq!(format!("{plain:?}"), format!("{fast:?}"));
    }

    #[test]
    fn cached_layout_matches_uncached_dynamic_list_and_result() {
        let ty = ValType::Record(vec![
            (
                "items".into(),
                ValType::List(Box::new(ValType::Record(vec![
                    ("id".into(), ValType::I32),
                    ("maybe".into(), ValType::Option(Box::new(ValType::I64))),
                ]))),
            ),
            (
                "outcome".into(),
                ValType::Result(None, Some(Box::new(ValType::String))),
            ),
        ]);
        let layout = crate::canon_layout::MemLayout::new(&ty);
        let item = |id: i32| {
            let mut object = crate::value::Object::new();
            object.properties.insert("id".into(), Value::I32(id));
            object
                .properties
                .insert("maybe".into(), Value::I64((id * 10) as i64));
            Value::Object(crate::heap::alloc(object))
        };
        let mut error = crate::value::Object::new();
        error
            .properties
            .insert("__wasi_error".into(), Value::String("denied".into()));
        let mut value = crate::value::Object::new();
        value.properties.insert(
            "items".into(),
            Value::Object(crate::heap::alloc(crate::value::Object::new_array(vec![
                item(2),
                item(4),
            ]))),
        );
        value
            .properties
            .insert("outcome".into(), Value::Object(crate::heap::alloc(error)));
        let value = Value::Object(crate::heap::alloc(value));
        let plain = SharedMemory::new(512);
        let cached = SharedMemory::new(512);
        let mut plain_next = 128u32;
        let mut cached_next = 128u32;
        let mut plain_alloc = |size: u32, align: u32| {
            let at = align_to(plain_next, align);
            plain_next = at + size;
            (plain_next <= 512).then_some(at)
        };
        let mut cached_alloc = |size: u32, align: u32| {
            let at = align_to(cached_next, align);
            cached_next = at + size;
            (cached_next <= 512).then_some(at)
        };
        store_with(&plain, &mut plain_alloc, &value, &ty, 0).unwrap();
        store_with_layout(&cached, &mut cached_alloc, &value, &ty, 0, &layout).unwrap();
        assert_eq!(plain_next, cached_next);
        let mut a = [0; 512];
        let mut b = [0; 512];
        plain.read_bytes(0, &mut a);
        cached.read_bytes(0, &mut b);
        assert_eq!(a, b);
        assert_eq!(
            format!("{:?}", load(&plain, &ty, 0).unwrap()),
            format!("{:?}", load_with_layout(&cached, &ty, 0, &layout).unwrap())
        );
    }

    #[test]
    #[ignore = "run explicitly for a controlled cached-vs-uncached comparison"]
    fn bench_cached_memory_layout() {
        let ty = ValType::Record(
            (0..12)
                .map(|index| {
                    (
                        format!("field{index}"),
                        ValType::Option(Box::new(ValType::I64)),
                    )
                })
                .collect(),
        );
        let layout = crate::canon_layout::MemLayout::new(&ty);
        let memory = SharedMemory::new(layout.size as usize);
        let mut value = crate::value::Object::new();
        for index in 0..12 {
            value
                .properties
                .insert(format!("field{index}").into(), Value::I64(index));
        }
        let value = Value::Object(crate::heap::alloc(value));
        let mut alloc = no_realloc;
        for (label, cached) in [("uncached", false), ("cached", true)] {
            let start = std::time::Instant::now();
            for _ in 0..20_000 {
                if cached {
                    store_with_layout(&memory, &mut alloc, &value, &ty, 0, &layout).unwrap();
                    std::hint::black_box(load_with_layout(&memory, &ty, 0, &layout).unwrap());
                } else {
                    store_with(&memory, &mut alloc, &value, &ty, 0).unwrap();
                    std::hint::black_box(load(&memory, &ty, 0).unwrap());
                }
            }
            eprintln!("canonical memory {label}: {:?}", start.elapsed());
        }
    }

    #[test]
    fn byte_list_round_trips_binary_and_narrowed_values() {
        let memory = SharedMemory::new(64);
        let value = Value::Object(crate::heap::alloc(crate::value::Object::new_array(vec![
            Value::I32(0),
            Value::I32(255),
            Value::I32(256),
            Value::I32(-1),
        ])));
        let mut alloc = |_size: u32, _align: u32| Some(32);
        let ty = ValType::List(Box::new(ValType::U8));

        store_with(&memory, &mut alloc, &value, &ty, 0).unwrap();
        assert_eq!(memory.load_i32(0).unwrap(), 32);
        assert_eq!(memory.load_i32(4).unwrap(), 4);
        assert_eq!(memory.read_array::<4>(32).unwrap(), [0, 255, 0, 255]);

        let back = load(&memory, &ty, 0).unwrap();
        assert_eq!(
            array_items(&back),
            vec![
                Value::I32(0),
                Value::I32(255),
                Value::I32(0),
                Value::I32(255),
            ]
        );
    }

    #[test]
    fn byte_list_store_rejects_out_of_bounds_payload() {
        let memory = SharedMemory::new(16);
        let value = Value::Object(crate::heap::alloc(crate::value::Object::new_array(vec![
            Value::I32(1),
            Value::I32(2),
        ])));
        let mut alloc = |_size: u32, _align: u32| Some(15);
        let ty = ValType::List(Box::new(ValType::U8));
        assert!(matches!(
            store_with(&memory, &mut alloc, &value, &ty, 0),
            Err(CanonError::OutOfBounds { .. })
        ));
    }

    #[test]
    fn byte_list_load_rejects_out_of_bounds_payload_before_allocating() {
        let memory = SharedMemory::new(16);
        memory.store_i32(0, 12).unwrap();
        memory.store_i32(4, 100).unwrap();
        let ty = ValType::List(Box::new(ValType::U8));
        assert!(matches!(
            load(&memory, &ty, 0),
            Err(CanonError::OutOfBounds { .. })
        ));
    }

    #[test]
    fn byte_list_load_enforces_cm3_maximum_and_empty_pointer_bounds() {
        let memory = SharedMemory::new(16);
        let ty = ValType::List(Box::new(ValType::U8));

        memory.store_i32(0, 0).unwrap();
        memory.store_i32(4, 1 << 28).unwrap();
        assert!(matches!(
            load(&memory, &ty, 0),
            Err(CanonError::ListByteLengthExceeded { .. })
        ));

        memory.store_i32(0, 17).unwrap();
        memory.store_i32(4, 0).unwrap();
        assert!(matches!(
            load(&memory, &ty, 0),
            Err(CanonError::OutOfBounds { .. })
        ));
    }

    #[test]
    fn byte_list_flat_pair_round_trips() {
        let memory = SharedMemory::new(64);
        let value = Value::Object(crate::heap::alloc(crate::value::Object::new_array(vec![
            Value::I32(0),
            Value::I32(128),
            Value::I32(255),
        ])));
        let ty = ValType::List(Box::new(ValType::U8));
        let mut alloc_fn = |_size: u32, _align: u32| Some(32);
        let mut alloc: Realloc<'_> = &mut alloc_fn;
        let (ptr, len) = store_pair_public(&memory, &mut alloc, &value, &ty).unwrap();
        assert_eq!((ptr, len), (32, 3));
        assert_eq!(memory.read_array::<3>(32).unwrap(), [0, 128, 255]);
        let back = load_pair_public(&memory, &ty, ptr, len).unwrap();
        assert_eq!(array_items(&back), array_items(&value));
    }

    #[test]
    #[ignore = "local canonical byte-list diagnostic"]
    fn bench_byte_list_store_and_load() {
        use std::time::Instant;

        const LEN: usize = 256 * 1024;
        let memory = SharedMemory::new(LEN + 64);
        let value = Value::Object(crate::heap::alloc(crate::value::Object::new_array(
            (0..LEN).map(|i| Value::I32((i & 255) as i32)).collect(),
        )));
        let ty = ValType::List(Box::new(ValType::U8));
        for _ in 0..3 {
            let start = Instant::now();
            let mut alloc = |_size: u32, _align: u32| Some(32);
            store_with(&memory, &mut alloc, &value, &ty, 0).unwrap();
            let store_time = start.elapsed();

            let start = Instant::now();
            std::hint::black_box(load(&memory, &ty, 0).unwrap());
            eprintln!(
                "canonical list<u8> {LEN} bytes: store {store_time:?}, load {:?}",
                start.elapsed()
            );
        }
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
            o.properties
                .insert("type".into(), Value::String(kind.into()));
            o.properties
                .insert("name".into(), Value::String(name.into()));
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
            Ok(Value::String(read_utf8_as(memory, at as usize, len as usize)?))
        }
        ValType::List(elem) => {
            if matches!(elem.as_ref(), ValType::U8) {
                return load_byte_items(memory, at, len);
            }
            let stride = elem_size(elem);
            if let Some(result) = load_scalar_items(memory, elem, at, len) {
                return result;
            }
            let mut items = Vec::with_capacity(len as usize);
            for i in 0..len {
                items.push(load(memory, elem, at + stride * i)?);
            }
            Ok(Value::Object(crate::heap::alloc(
                crate::value::Object::new_array(items),
            )))
        }
        _ => Err(CanonError::Unsupported(
            "load_pair: not a (ptr, length) type",
        )),
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
            let text = string_text(v);
            let bytes = text.as_bytes();
            let size = bytes.len() as u32;
            let at = realloc(size.max(1), 1).ok_or(CanonError::OutOfMemory { size })?;
            write_bytes(memory, at, bytes)?;
            Ok((at, size))
        }
        ValType::List(elem) => {
            if matches!(elem.as_ref(), ValType::U8) {
                let bytes = byte_items(v)?;
                let size = bytes.len() as u32;
                let at = realloc(size.max(1), 1).ok_or(CanonError::OutOfMemory { size })?;
                write_bytes(memory, at, &bytes)?;
                return Ok((at, size));
            }
            let items = array_items(v);
            let stride = elem_size(elem);
            if let Some(result) = store_scalar_list(memory, &mut **realloc, elem, &items, true) {
                return result;
            }
            let align = alignment(elem);
            let size = stride * items.len() as u32;
            let at = realloc(size.max(1), align).ok_or(CanonError::OutOfMemory { size })?;
            for (i, item) in items.iter().enumerate() {
                store_with(memory, realloc, item, elem, at + stride * i as u32)?;
            }
            Ok((at, items.len() as u32))
        }
        _ => Err(CanonError::Unsupported(
            "store_pair: not a (ptr, length) type",
        )),
    }
}

/// Two bytes, little-endian — the memory has no 16-bit primitive.
fn load_u16(memory: &crate::shared_memory::SharedMemory, addr: usize) -> u16 {
    memory
        .read_array::<2>(addr)
        .map(u16::from_le_bytes)
        .unwrap_or(0)
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
        object.properties.insert(
            label.as_str().into(),
            Value::Bool((packed >> shift) & 1 == 1),
        );
    }
    Value::Object(crate::heap::alloc(object))
}

#[cfg(test)]
mod scalar_list_bulk_tests {
    use super::*;
    use crate::shared_memory::SharedMemory;

    fn assert_same(actual: &[Value], expected: &[Value]) {
        assert_eq!(actual.len(), expected.len());
        for (a, b) in actual.iter().zip(expected) {
            match (a, b) {
                (Value::F64(a), Value::F64(b)) => assert_eq!(a.to_bits(), b.to_bits()),
                _ => assert_eq!(a, b),
            }
        }
    }

    #[test]
    fn scalar_lists_match_recursive_lifting_for_all_widths() {
        let memory = SharedMemory::from_vec((0..512).map(|i| (i * 73 + 129) as u8).collect());
        for elem in [ValType::Bool, ValType::S8, ValType::U8, ValType::S16,
                     ValType::U16, ValType::I32, ValType::U32, ValType::I64,
                     ValType::U64, ValType::F32, ValType::F64] {
            let count = 32;
            let stride = elem_size(&elem);
            let expected: Vec<_> = (0..count).map(|i| load(&memory, &elem, 16 + i * stride).unwrap()).collect();
            let ty = ValType::List(Box::new(elem));
            assert_same(&array_items(&load_pair_public(&memory, &ty, 16, count).unwrap()), &expected);
            store_pair_header(&memory, 0, 16, count);
            assert_same(&array_items(&load(&memory, &ty, 0).unwrap()), &expected);
        }
    }

    #[test]
    fn scalar_list_ranges_reject_invalid_lengths_before_allocating() {
        let memory = SharedMemory::new(64);
        let ty = ValType::List(Box::new(ValType::I64));
        for (ptr, count) in [(56, 2), (72, 0), (0, u32::MAX)] {
            assert!(load_pair_public(&memory, &ty, ptr, count).is_err());
        }
        assert!(matches!(load_pair_public(&memory, &ty, 1, 0), Err(CanonError::Misaligned { .. })));
        assert!(matches!(load_pair_public(&memory, &ty, 0, 1 << 25), Err(CanonError::ListByteLengthExceeded { .. })));
        assert!(array_items(&load_pair_public(&memory, &ty, 64, 0).unwrap()).is_empty());
    }

    #[test]
    fn scalar_lists_preserve_float_bits_and_own_results() {
        let memory = SharedMemory::new(64);
        let bits = [0x8000000000000000u64, 0x7ff0000000000000, 0x7ff8000000000123];
        for (i, bits) in bits.iter().enumerate() { memory.write_array(i * 8, &bits.to_le_bytes()).unwrap(); }
        let ty = ValType::List(Box::new(ValType::F64));
        let result = load_pair_public(&memory, &ty, 0, 3).unwrap();
        memory.write_array(0, &[0u8; 24]).unwrap();
        for (value, bits) in array_items(&result).iter().zip(bits) {
            match value { Value::F64(value) => assert_eq!(value.to_bits(), bits), _ => panic!("wrong scalar type") }
        }
    }

    #[test]
    #[ignore = "paired loaded-machine scalar list memory measurement"]
    fn scalar_list_bulk_loaded_benchmark() {
        use std::hint::black_box;
        use std::time::Instant;
        for elem in [ValType::I32, ValType::I64, ValType::F64] {
            for count in [64u32, 4096, 32768] {
                let stride = elem_size(&elem);
                let memory = SharedMemory::from_vec((0..(count * stride) as usize).map(|i| i as u8).collect());
                let ty = ValType::List(Box::new(elem.clone()));
                let iterations = (65536 / count).max(4);
                let mut old_times = Vec::new();
                let mut new_times = Vec::new();
                for pair in 0..23 {
                    for new in if pair % 2 == 0 { [false, true] } else { [true, false] } {
                        let start = Instant::now();
                        for _ in 0..iterations {
                            let result = if new {
                                load_pair_public(&memory, black_box(&ty), 0, count).unwrap()
                            } else {
                                let mut items = Vec::with_capacity(count as usize);
                                for i in 0..count { items.push(load(&memory, black_box(&elem), i * stride).unwrap()); }
                                Value::Object(crate::heap::alloc(crate::value::Object::new_array(items)))
                            };
                            black_box(result);
                        }
                        let duration = start.elapsed().as_nanos();
                        if pair >= 2 { if new { new_times.push(duration); } else { old_times.push(duration); } }
                    }
                }
                old_times.sort_unstable(); new_times.sort_unstable();
                eprintln!("scalar-list elem={elem:?} count={count} old_p50_ns={} new_p50_ns={} old_p95_ns={} new_p95_ns={} ratio={:.3}",
                    old_times[10], new_times[10], old_times[19], new_times[19], old_times[10] as f64 / new_times[10] as f64);
            }
        }
    }
}

#[cfg(test)]
mod plain_record_list_tests {
    use super::*;
    use crate::canon_layout::MemLayout;
    use crate::shared_memory::SharedMemory;

    #[test]
    fn cloned_record_maps_preserve_order_values_and_independence() {
        for names in [
            vec!["left", "right"],
            vec!["same", "same", "tail"],
            vec!["a", "b", "c", "d", "e", "f", "g", "h"],
        ] {
            let fields: Vec<_> = names.iter().map(|name| ((*name).into(), ValType::I32)).collect();
            let parts: Vec<_> = (0..fields.len())
                .map(|index| (index as u32 * 4, MemLayout::new(&ValType::I32)))
                .collect();
            let size = fields.len() * 4;
            for count in [0, 1, 64] {
                let bytes: Vec<_> = (0..count * fields.len())
                    .flat_map(|value| (value as i32).to_le_bytes())
                    .collect();
                let expected = array_items(&decode_plain_record_items_with_capacity::<true>(
                    &fields, &parts, size, &bytes,
                ));
                let actual = array_items(&decode_plain_record_items(&fields, &parts, size, &bytes));
                assert_eq!(actual.len(), count);
                for (actual, expected) in actual.iter().zip(&expected) {
                    let (Value::Object(actual), Value::Object(expected)) = (actual, expected) else {
                        panic!("record list must contain objects");
                    };
                    let actual = actual.lock().unwrap();
                    let expected = expected.lock().unwrap();
                    assert_eq!(actual.properties, expected.properties);
                    assert_eq!(actual.properties.keys().collect::<Vec<_>>(),
                        expected.properties.keys().collect::<Vec<_>>());
                }
                if count > 1 {
                    let Value::Object(first) = &actual[0] else { unreachable!() };
                    let Value::Object(second) = &actual[1] else { unreachable!() };
                    let previous = second.lock().unwrap().properties[names[0]].clone();
                    first.lock().unwrap().properties.insert(names[0].into(), Value::I32(-1));
                    assert_eq!(second.lock().unwrap().properties[names[0]], previous);
                }
            }
        }
    }

    #[test]
    fn cached_plain_records_match_recursive_values_and_traps() {
        let fields = vec![
            ("bool", ValType::Bool, Value::Bool(true)),
            ("s8", ValType::S8, Value::I32(-3)),
            ("u8", ValType::U8, Value::I32(253)),
            ("s16", ValType::S16, Value::I32(-1234)),
            ("u16", ValType::U16, Value::I32(60_000)),
            ("i32", ValType::I32, Value::I32(-12_345)),
            ("u32", ValType::U32, Value::F64(u32::MAX as f64)),
            ("i64", ValType::I64, Value::I64(-12_345_678_901)),
            ("u64", ValType::U64, Value::bigint_u64(u64::MAX)),
            ("f32", ValType::F32, Value::F64(-0.0)),
            ("f64", ValType::F64, Value::F64(1.25)),
        ];
        let elem = ValType::Record(
            fields.iter().map(|(name, ty, _)| (name.to_string(), ty.clone())).collect(),
        );
        let ty = ValType::List(Box::new(elem.clone()));
        let layout = MemLayout::new(&ty);
        let stride = elem_size(&elem);
        let at = 64u32;
        let count = 3u32;
        let memory = SharedMemory::new((at + count * stride) as usize);
        let mut value = crate::value::Object::new();
        for (name, _, field_value) in &fields {
            value.properties.insert((*name).into(), field_value.clone());
        }
        let value = Value::Object(crate::heap::alloc(value));
        for index in 0..count {
            store(&memory, &value, &elem, at + index * stride).unwrap();
        }
        store_pair_header(&memory, 0, at, count);
        let plain = array_items(&load(&memory, &ty, 0).unwrap());
        let cached = array_items(&load_with_layout(&memory, &ty, 0, &layout).unwrap());
        assert_eq!(plain.len(), cached.len());
        for (plain, cached) in plain.iter().zip(&cached) {
            let (Value::Object(plain), Value::Object(cached)) = (plain, cached) else {
                panic!("record list must contain objects");
            };
            let plain = plain.lock().unwrap();
            let cached = cached.lock().unwrap();
            for (name, _, _) in &fields {
                assert_eq!(plain.properties.get(*name), cached.properties.get(*name), "{name}");
            }
            assert_eq!(
                plain.properties["f32"].as_f64().to_bits(),
                cached.properties["f32"].as_f64().to_bits()
            );
        }

        for (ptr, len) in [(at, count + 1), (at + 1, count)] {
            store_pair_header(&memory, 0, ptr, len);
            let plain = load(&memory, &ty, 0).unwrap_err();
            let cached = load_with_layout(&memory, &ty, 0, &layout).unwrap_err();
            assert_eq!(format!("{plain:?}"), format!("{cached:?}"));
        }

        let char_elem = ValType::Record(vec![("char".into(), ValType::Char)]);
        let char_ty = ValType::List(Box::new(char_elem));
        let char_layout = MemLayout::new(&char_ty);
        store_pair_header(&memory, 0, at, 1);
        memory.write_array(at as usize, &0xd800u32.to_le_bytes()).unwrap();
        let plain = load(&memory, &char_ty, 0).unwrap_err();
        let cached = load_with_layout(&memory, &char_ty, 0, &char_layout).unwrap_err();
        assert_eq!(format!("{plain:?}"), format!("{cached:?}"));
    }

    #[test]
    #[ignore = "paired canonical record-list decode benchmark"]
    fn bench_cached_plain_records_against_recursive_decoder() {
        use std::hint::black_box;
        use std::time::Instant;

        let elem = ValType::Record(vec![
            ("left".into(), ValType::I32),
            ("right".into(), ValType::I32),
        ]);
        let ty = ValType::List(Box::new(elem.clone()));
        let layout = MemLayout::new(&ty);
        let crate::canon_layout::MemShape::Elements(child) = &layout.shape else {
            unreachable!();
        };
        let crate::canon_layout::MemShape::Record(parts) = &child.shape else {
            unreachable!();
        };
        let ValType::Record(fields) = &elem else {
            unreachable!();
        };
        let stride = child.size;
        for count in [64u32, 4096, 32768] {
            let at = 64u32;
            let memory = SharedMemory::new((at + count * stride) as usize);
            store_pair_header(&memory, 0, at, count);
            memory.map_write_range(at as usize, (count * stride) as usize, |range| {
                for (index, record) in range.chunks_exact_mut(stride as usize).enumerate() {
                    record[..4].copy_from_slice(&(index as i32).to_le_bytes());
                    record[4..8].copy_from_slice(&(!(index as i32)).to_le_bytes());
                }
            }).unwrap();
            let iterations = (32768 / count).max(2);
            let mut recursive = Vec::new();
            let mut bulk = Vec::new();
            let mut locked = Vec::new();
            for round in 0..11 {
                for mode in if round % 2 == 0 { [0, 1, 2] } else { [2, 1, 0] } {
                    let start = Instant::now();
                    for _ in 0..iterations {
                        let result = match mode {
                            0 => load_list_items_recursive(&memory, &elem, at, count, stride, Some(child)).unwrap(),
                            1 => load_with_layout(&memory, &ty, 0, &layout).unwrap(),
                            _ => memory.map_read_range(at as usize, (count * stride) as usize, |bytes| {
                                decode_plain_record_items(fields, parts, stride as usize, bytes)
                            }).unwrap(),
                        };
                        black_box(result);
                    }
                    if round > 1 {
                        (match mode { 0 => &mut recursive, 1 => &mut bulk, _ => &mut locked })
                            .push(start.elapsed().as_nanos() / iterations as u128);
                    }
                }
            }
            recursive.sort_unstable();
            bulk.sort_unstable();
            locked.sort_unstable();
            eprintln!(
                "record-list count={count}: recursive={} ns, snapshot={} ns, locked={} ns",
                recursive[recursive.len() / 2],
                bulk[bulk.len() / 2],
                locked[locked.len() / 2],
            );
        }
    }

    #[test]
    #[ignore = "paired canonical record map-template benchmark"]
    fn bench_plain_record_map_template() {
        use std::hint::black_box;
        use std::time::Instant;

        for width in [1usize, 2, 4, 8] {
            let fields: Vec<_> = (0..width)
                .map(|index| (format!("field{index}"), ValType::I32)).collect();
            let parts: Vec<_> = (0..width)
                .map(|index| (index as u32 * 4, MemLayout::new(&ValType::I32))).collect();
            for count in [1usize, 8, 64, 4096, 32768] {
                let bytes: Vec<_> = (0..count * width)
                    .flat_map(|value| (value as i32).to_le_bytes()).collect();
                let iterations = (32768 / count).max(2);
                let mut old = Vec::new();
                let mut new = Vec::new();
                for round in 0..9 {
                    for template in if round % 2 == 0 { [false, true] } else { [true, false] } {
                        let start = Instant::now();
                        for _ in 0..iterations {
                            black_box(if template {
                                decode_plain_record_items(&fields, &parts, width * 4, &bytes)
                            } else {
                                decode_plain_record_items_with_capacity::<true>(
                                    &fields, &parts, width * 4, &bytes,
                                )
                            });
                        }
                        if round != 0 {
                            (if template { &mut new } else { &mut old })
                                .push(start.elapsed().as_nanos() / iterations as u128);
                        }
                    }
                }
                old.sort_unstable();
                new.sort_unstable();
                eprintln!("record-map fields={width} count={count}: insert={} ns, template={} ns",
                    old[old.len() / 2], new[new.len() / 2]);
            }
        }
    }

    #[test]
    #[ignore = "paired canonical record field-capacity benchmark"]
    fn bench_plain_record_capacity() {
        use std::hint::black_box;
        use std::time::Instant;

        for field_count in [2usize, 4, 8] {
            let fields: Vec<_> = (0..field_count)
                .map(|index| (format!("field{index}"), ValType::I32))
                .collect();
            let layout = MemLayout::new(&ValType::Record(fields.clone()));
            let crate::canon_layout::MemShape::Record(parts) = &layout.shape else {
                unreachable!();
            };
            let size = layout.size as usize;
            for count in [64usize, 4096, 32768] {
                let mut bytes = vec![0; count * size];
                for (index, record) in bytes.chunks_exact_mut(size).enumerate() {
                    for (field, (offset, _)) in parts.iter().enumerate() {
                        let at = *offset as usize;
                        record[at..at + 4]
                            .copy_from_slice(&((index ^ field) as i32).to_le_bytes());
                    }
                }
                let iterations = (32768 / count).max(1);
                let mut reserved = Vec::new();
                let mut unreserved = Vec::new();
                for round in 0..11 {
                    for mode in if round % 2 == 0 { [true, false] } else { [false, true] } {
                        let start = Instant::now();
                        for _ in 0..iterations {
                            let value = if mode {
                                decode_plain_record_items_with_capacity::<true>(
                                    &fields, parts, size, &bytes,
                                )
                            } else {
                                decode_plain_record_items_with_capacity::<false>(
                                    &fields, parts, size, &bytes,
                                )
                            };
                            black_box(value);
                        }
                        if round > 1 {
                            (if mode { &mut reserved } else { &mut unreserved })
                                .push(start.elapsed().as_nanos() / iterations as u128);
                        }
                    }
                }
                reserved.sort_unstable();
                unreserved.sort_unstable();
                eprintln!(
                    "record fields={field_count} count={count}: reserved={} ns, unreserved={} ns",
                    reserved[reserved.len() / 2],
                    unreserved[unreserved.len() / 2],
                );
            }
        }
    }
}

#[cfg(test)]
mod scalar_list_store_tests {
    use super::*;
    use crate::shared_memory::SharedMemory;

    #[test]
    fn unsigned_lists_keep_fast_path_and_exact_bits() {
        for (elem, items) in [
            (ValType::U32, vec![Value::F64(0.0), Value::F64(2_147_483_648.0), Value::F64(4_294_967_295.0)]),
            (ValType::U64, vec![Value::bigint_u64(0), Value::bigint_u64(1 << 63), Value::bigint_u64(u64::MAX)]),
        ] {
            let memory = SharedMemory::new(64);
            let value = Value::Object(crate::heap::alloc(crate::value::Object::new_array(items.clone())));
            let mut allocate = |_, _| Some(16);
            let mut realloc: Realloc<'_> = &mut allocate;
            assert_eq!(store_pair_public(&memory, &mut realloc, &value, &ValType::List(Box::new(elem.clone()))).unwrap(), (16, 3));
            for (i, original) in items.iter().enumerate() {
                let loaded = load(&memory, &elem, 16 + i as u32 * elem_size(&elem)).unwrap();
                match (original, loaded) {
                    (Value::F64(want), Value::F64(got)) => assert_eq!(*want, got),
                    (Value::BigInt(want), Value::BigInt(got)) => assert_eq!(want.to_u64_wrapping(), got.to_u64_wrapping()),
                    other => panic!("unexpected unsigned list value: {other:?}"),
                }
            }
        }
    }

    #[test]
    fn invalid_unsigned_list_rejects_before_realloc_or_memory_write() {
        let memory = SharedMemory::new(32);
        memory.write_array(0, &[0xa5; 32]).unwrap();
        let value = Value::Object(crate::heap::alloc(crate::value::Object::new_array(vec![
            Value::F64(1.0),
            Value::F64(4_294_967_296.0),
        ])));
        let mut called = false;
        let mut allocate = |_, _| { called = true; Some(16) };
        let mut realloc: Realloc<'_> = &mut allocate;
        assert!(matches!(
            store_pair_public(&memory, &mut realloc, &value, &ValType::List(Box::new(ValType::U32))),
            Err(CanonError::UnsignedOutOfRange { bits: 32 })
        ));
        assert!(!called);
        assert_eq!(memory.read_array::<32>(0).unwrap(), [0xa5; 32]);
    }

    #[test]
    fn bulk_scalar_stores_match_recursive_bytes_and_both_shapes() {
        let items = vec![Value::Null, Value::Bool(true), Value::I32(-1),
            Value::I64(i64::MIN), Value::F64(-0.0), Value::F64(f64::INFINITY)];
        let value = Value::Object(crate::heap::alloc(crate::value::Object::new_array(items.clone())));
        for elem in [ValType::Bool, ValType::S8, ValType::S16, ValType::U16,
                     ValType::I32, ValType::I64, ValType::F32, ValType::F64] {
            let expected = SharedMemory::new(128);
            let actual = SharedMemory::new(128);
            let stride = elem_size(&elem);
            let mut allocate = |_, _| Some(32);
            for (i, item) in items.iter().enumerate() {
                store_with(&expected, &mut allocate, item, &elem, 32 + i as u32 * stride).unwrap();
            }
            let ty = ValType::List(Box::new(elem));
            let mut realloc: Realloc<'_> = &mut allocate;
            assert_eq!(store_pair_public(&actual, &mut realloc, &value, &ty).unwrap(), (32, items.len() as u32));
            assert_eq!(actual.read_array::<96>(32).unwrap(), expected.read_array::<96>(32).unwrap());
            store_with(&actual, &mut allocate, &value, &ty, 0).unwrap();
            assert_eq!(load_pair_header(&actual, 0), (32, items.len() as u32));
            assert_eq!(actual.read_array::<96>(32).unwrap(), expected.read_array::<96>(32).unwrap());
        }
    }

    #[test]
    fn source_snapshot_and_realloc_memory_relocation_are_preserved() {
        let memory = SharedMemory::new(16);
        let value = Value::Object(crate::heap::alloc(crate::value::Object::new_array(vec![Value::I32(7), Value::I32(-3)])));
        let mut allocate = |size, align| {
            assert_eq!((size, align), (8, 4));
            memory.with_buffer_mut(|bytes| bytes.resize(96, 0));
            if let Value::Object(object) = &value {
                if let crate::value::ObjectKind::Array(items) = &mut object.lock().unwrap().kind {
                    items[0] = Value::I32(99);
                }
            }
            Some(32)
        };
        store_with(&memory, &mut allocate, &value, &ValType::List(Box::new(ValType::I32)), 0).unwrap();
        assert_eq!(memory.load_i32(32).unwrap(), 7);
        assert_eq!(memory.load_i32(36).unwrap(), -3);
        assert_eq!(array_items(&value)[0], Value::I32(99));
    }

    #[test]
    fn invalid_allocator_ranges_trap_before_mutating_and_empty_requests_match() {
        let memory = SharedMemory::from_vec(vec![0x5a; 64]);
        let value = Value::Object(crate::heap::alloc(crate::value::Object::new_array(vec![Value::I64(1), Value::I64(2)])));
        let ty = ValType::List(Box::new(ValType::I64));
        for dest in [1, 56] {
            let mut allocate = |_, _| Some(dest);
            assert!(store_with(&memory, &mut allocate, &value, &ty, 0).is_err());
            assert_eq!(memory.read_array::<64>(0).unwrap(), [0x5a; 64]);
        }
        let empty = Value::Object(crate::heap::alloc(crate::value::Object::new_array(Vec::new())));
        let mut called = false;
        let mut allocate = |_, _| { called = true; Some(32) };
        store_with(&memory, &mut allocate, &empty, &ty, 0).unwrap();
        assert!(!called);
        let mut request = None;
        let mut allocate = |size, align| { request = Some((size, align)); Some(32) };
        let mut realloc: Realloc<'_> = &mut allocate;
        assert_eq!(store_pair_public(&memory, &mut realloc, &empty, &ty).unwrap(), (32, 0));
        assert_eq!(request, Some((1, 8)));
    }

    #[test]
    fn nonprimitive_coercions_remain_on_recursive_path() {
        let memory = SharedMemory::new(64);
        let values = vec![Value::String("12".into()), Value::String("-7".into())];
        let expected = SharedMemory::new(64);
        let mut allocate = |_, _| Some(32);
        for (i, value) in values.iter().enumerate() {
            store_with(&expected, &mut allocate, value, &ValType::I32, 32 + i as u32 * 4).unwrap();
        }
        let value = Value::Object(crate::heap::alloc(crate::value::Object::new_array(values)));
        let mut realloc: Realloc<'_> = &mut allocate;
        store_pair_public(&memory, &mut realloc, &value, &ValType::List(Box::new(ValType::I32))).unwrap();
        assert_eq!(memory.read_array::<8>(32).unwrap(), expected.read_array::<8>(32).unwrap());
    }

    #[test]
    #[ignore = "paired loaded-machine scalar-list write measurement"]
    fn scalar_list_store_loaded_benchmark() {
        use std::hint::black_box;
        use std::time::Instant;
        for elem in [ValType::I32, ValType::I64, ValType::F64] {
            for count in [64u32, 4096, 32768] {
                let stride = elem_size(&elem);
                let memory = SharedMemory::new((count * stride + 32) as usize);
                let value = Value::Object(crate::heap::alloc(crate::value::Object::new_array(
                    (0..count).map(|i| Value::I32(i as i32)).collect())));
                let ty = ValType::List(Box::new(elem.clone()));
                let iterations = (65536 / count).max(4);
                let mut old_times = Vec::new();
                let mut new_times = Vec::new();
                for pair in 0..23 {
                    for new in if pair % 2 == 0 { [false, true] } else { [true, false] } {
                        let start = Instant::now();
                        for _ in 0..iterations {
                            let mut allocate = |_, _| Some(32);
                            if new {
                                let mut realloc: Realloc<'_> = &mut allocate;
                                black_box(store_pair_public(&memory, &mut realloc, black_box(&value), black_box(&ty)).unwrap());
                            } else {
                                let items = array_items(black_box(&value));
                                let dest = allocate((stride * items.len() as u32).max(1), alignment(&elem)).unwrap();
                                for (i, item) in items.iter().enumerate() {
                                    store_with(&memory, &mut allocate, item, black_box(&elem), dest + i as u32 * stride).unwrap();
                                }
                                black_box((dest, items.len() as u32));
                            }
                        }
                        let duration = start.elapsed().as_nanos();
                        if pair >= 2 { if new { new_times.push(duration); } else { old_times.push(duration); } }
                    }
                }
                old_times.sort_unstable(); new_times.sort_unstable();
                eprintln!("scalar-store elem={elem:?} count={count} old_p50_ns={} new_p50_ns={} old_p95_ns={} new_p95_ns={} ratio={:.3}",
                    old_times[10], new_times[10], old_times[19], new_times[19], old_times[10] as f64 / new_times[10] as f64);
            }
        }
    }
}
