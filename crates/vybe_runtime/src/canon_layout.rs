//! Canonical ABI memory layout — `CanonicalABI.md` §`alignment` / §`elem_size`.
//!
//! How many bytes one value of a component type occupies in linear memory, and
//! what it must be aligned to. This is the half of §Storing/§Loading that has
//! to exist before anything can copy a typed value in or out:
//!
//! - `canon future.{read,write}` is `(func (param i32 T) (result i32))` — a
//!   handle and a POINTER, with **no count**, because a future carries exactly
//!   one element whose size comes from the type. Without a size there is no
//!   way to know how many bytes to move, which is why `future.read` could not
//!   be implemented at all before this.
//! - `canon stream.{read,write}` takes a count of ELEMENTS, not bytes.
//!   `stream<u8>` happens to make those the same number, which is the only
//!   reason byte streams worked without any of this.
//!
//! The rules are transcribed rather than invented; each function names the
//! spec function it mirrors so the two can be diffed. Where the spec threads a
//! `ptr_type` for the 🐘 memory64 gate, this assumes a 32-bit address space —
//! the one place a wider address type would change an answer is marked.

use crate::component::ValType;

/// `ptr_size(ptr_type)` for a 32-bit address space.
///
/// 🐘 memory64 would make this 8 and widen `string` and unfixed `list` with
/// it. Every use goes through this constant so that becomes one edit rather
/// than a hunt.
pub(crate) const PTR_SIZE: u32 = 4;

/// `align_to(ptr, alignment)` — round `offset` up to the next multiple.
pub fn align_to(offset: u32, alignment: u32) -> u32 {
    debug_assert!(alignment.is_power_of_two());
    offset.div_ceil(alignment) * alignment
}

/// `alignment(t)` — `CanonicalABI.md:2225`.
pub fn alignment(t: &ValType) -> u32 {
    match t {
        ValType::Bool => 1,
        // `CanonicalABI.md:2227` — the narrow widths align to their OWN size,
        // not to 4. Aligning an `s8` field to 4 would put every later field of
        // a record at the wrong offset.
        ValType::S8 | ValType::U8 => 1,
        ValType::S16 | ValType::U16 => 2,
        ValType::I32 => 4,
        ValType::I64 => 8,
        ValType::F32 => 4,
        ValType::F64 => 8,
        // A Unicode scalar value is carried as four bytes.
        ValType::Char => 4,
        // `alignment_flags` — the packed integer's own width.
        ValType::Flags(labels) => flags_bytes(labels.len()),
        // `string` is (ptr, length) — aligned as a pointer.
        ValType::String => PTR_SIZE,
        // An unfixed `list` is also (ptr, length).
        ValType::List(_) => PTR_SIZE,
        // 🔧 `alignment_list(t, N)` with a length present is the ELEMENT's
        // alignment — the elements are inline, so there is no pointer to align.
        ValType::ListFixed(elem, _) => alignment(elem),
        ValType::Record(fields) => fields.iter().map(|(_, t)| alignment(t)).max().unwrap_or(1),
        // `option`/`result` despecialise to `variant`, so they align as one:
        // the max of the discriminant's alignment and every case payload's.
        ValType::Option(inner) => alignment_variant(&[None, Some(inner.as_ref())]),
        // `result` despecialises to a two-case variant, and EITHER case may
        // be payload-free — `result<_, error-code>` is the common one.
        ValType::Result(ok, err) => alignment_variant(&[ok.as_deref(), err.as_deref()]),
        ValType::Variant(cases) => alignment_variant(&variant_cases(cases)),
        // Handles are i32 indices into the handle table.
        ValType::Own(_) | ValType::Borrow(_) | ValType::ErrorContext => 4,
        ValType::Stream(_) | ValType::Future(_) => 4,
        // Not a component type — see `elem_size`.
        ValType::Any => 1,
    }
}

/// Borrow a `variant`'s cases in the `&[Option<&ValType>]` shape the spec
/// helpers take — they only ever need each case's payload type, never its name.
pub(crate) fn variant_cases(cases: &[(String, Option<ValType>)]) -> Vec<Option<&ValType>> {
    cases.iter().map(|(_, t)| t.as_ref()).collect()
}

/// Byte width of a `variant`'s discriminant — `CanonicalABI.md` §`discriminant_type`.
///
/// Public because §`store_variant` writes the discriminant itself and then has
/// to know where the payload starts; both answers come from here so they cannot
/// drift apart.
pub fn variant_discriminant_size(cases: &[(String, Option<ValType>)]) -> u32 {
    discriminant_size(cases.len())
}

/// Offset of a `variant`'s payload — `CanonicalABI.md` §`store_variant`:
/// `payload_ptr = ptr + align_to(discriminant_size, max_case_alignment)`.
pub fn variant_payload_offset(cases: &[(String, Option<ValType>)]) -> u32 {
    let borrowed = variant_cases(cases);
    align_to(discriminant_size(cases.len()), max_case_alignment(&borrowed))
}

/// `alignment_variant(cases)` — `CanonicalABI.md:2269`.
fn alignment_variant(cases: &[Option<&ValType>]) -> u32 {
    alignment(&discriminant_type(cases.len())).max(max_case_alignment(cases))
}

/// `max_case_alignment(cases)` — `CanonicalABI.md:2281`.
fn max_case_alignment(cases: &[Option<&ValType>]) -> u32 {
    cases
        .iter()
        .filter_map(|c| c.map(alignment))
        .max()
        .unwrap_or(1)
}

/// `discriminant_type(cases)` — `CanonicalABI.md:2272`. The SMALLEST integer
/// that covers the case count, which is what lets a two-case `result` pack
/// into fewer bytes than a naive tag would.
fn discriminant_type(case_count: usize) -> ValType {
    debug_assert!(case_count > 0);
    // ceil(log2(n) / 8): 0 or 1 → u8, 2 → u16, 3 → u32.
    match (case_count as f64).log2().ceil() as u32 / 8 {
        0 | 1 => ValType::Bool, // one byte, like the spec's U8Type
        2 => ValType::I32,      // widened below by elem_size_discriminant
        _ => ValType::I32,
    }
}

/// Byte width of the discriminant itself, kept separate from `discriminant_type`
/// because this crate's `ValType` has no `u8`/`u16` to name.
fn discriminant_size(case_count: usize) -> u32 {
    debug_assert!(case_count > 0);
    match (case_count as f64).log2().ceil() as u32 / 8 {
        0 | 1 => 1,
        2 => 2,
        _ => 4,
    }
}

/// `elem_size(t)` — `CanonicalABI.md:2311`. Bytes one value occupies.
pub fn elem_size(t: &ValType) -> u32 {
    match t {
        ValType::Bool => 1,
        ValType::S8 | ValType::U8 => 1,
        ValType::S16 | ValType::U16 => 2,
        ValType::I32 => 4,
        ValType::I64 => 8,
        ValType::F32 => 4,
        ValType::F64 => 8,
        ValType::Char => 4,
        // `elem_size_flags` is the SAME function as `alignment_flags` — one
        // packed integer, so its size and its alignment are one number.
        ValType::Flags(labels) => flags_bytes(labels.len()),
        // `string` and an unfixed `list` are both (ptr, length) pairs.
        ValType::String => 2 * PTR_SIZE,
        ValType::List(_) => 2 * PTR_SIZE,
        // 🔧 `N * elem_size(t)` — the whole list occupies the value's own
        // space, not a pointer's.
        ValType::ListFixed(elem, n) => n.saturating_mul(elem_size(elem)),
        ValType::Record(fields) => elem_size_record(fields),
        ValType::Option(inner) => elem_size_variant(&[None, Some(inner.as_ref())]),
        // `result` despecialises to a two-case variant, and EITHER case may
        // be payload-free — `result<_, error-code>` is the common one.
        ValType::Result(ok, err) => elem_size_variant(&[ok.as_deref(), err.as_deref()]),
        ValType::Variant(cases) => elem_size_variant(&variant_cases(cases)),
        ValType::Own(_) | ValType::Borrow(_) | ValType::ErrorContext => 4,
        ValType::Stream(_) | ValType::Future(_) => 4,
        // `Any` is this crate's escape hatch for dynamically typed frontends,
        // not a component type. It has no canonical layout; treating it as one
        // byte is a placeholder, and a caller that needs a real size should
        // reject it rather than silently move the wrong number of bytes.
        ValType::Any => 1,
    }
}

/// `elem_size_record(fields)` — `CanonicalABI.md:2335`. Each field is aligned
/// before it is placed, and the whole record is padded to its own alignment.
fn elem_size_record(fields: &[(String, ValType)]) -> u32 {
    let mut s = 0u32;
    for (_, t) in fields {
        s = align_to(s, alignment(t));
        s += elem_size(t);
    }
    let a = fields.iter().map(|(_, t)| alignment(t)).max().unwrap_or(1);
    align_to(s, a)
}

/// `elem_size_variant(cases)` — `CanonicalABI.md:2346`. Discriminant, padded to
/// the widest case alignment, then the LARGEST payload — the cases overlap.
fn elem_size_variant(cases: &[Option<&ValType>]) -> u32 {
    let mut s = discriminant_size(cases.len());
    s = align_to(s, max_case_alignment(cases));
    let payload = cases
        .iter()
        .filter_map(|c| c.map(elem_size))
        .max()
        .unwrap_or(0);
    s += payload;
    align_to(s, alignment_variant(cases))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn scalars_match_the_spec_table() {
        assert_eq!((alignment(&ValType::Bool), elem_size(&ValType::Bool)), (1, 1));
        assert_eq!((alignment(&ValType::I32), elem_size(&ValType::I32)), (4, 4));
        assert_eq!((alignment(&ValType::I64), elem_size(&ValType::I64)), (8, 8));
        assert_eq!((alignment(&ValType::F64), elem_size(&ValType::F64)), (8, 8));
    }

    #[test]
    fn string_and_list_are_pointer_length_pairs() {
        assert_eq!(elem_size(&ValType::String), 8);
        assert_eq!(alignment(&ValType::String), 4);
        let l = ValType::List(Box::new(ValType::I64));
        assert_eq!(elem_size(&l), 8, "an unfixed list is (ptr, len), not its payload");
        assert_eq!(alignment(&l), 4);
    }

    #[test]
    fn handles_stream_and_future_are_i32_indices() {
        for t in [
            ValType::Stream(Box::new(ValType::Bool)),
            ValType::Future(Box::new(ValType::Bool)),
            ValType::Own("node".into()),
            ValType::Borrow("node".into()),
        ] {
            assert_eq!((alignment(&t), elem_size(&t)), (4, 4));
        }
    }

    #[test]
    fn record_pads_between_fields_and_at_the_end() {
        // (bool, i64): 1 byte, padded to 8, + 8 = 16, already 8-aligned.
        let r = ValType::Record(vec![
            ("flag".into(), ValType::Bool),
            ("n".into(), ValType::I64),
        ]);
        assert_eq!(alignment(&r), 8);
        assert_eq!(elem_size(&r), 16);
    }

    #[test]
    fn result_is_a_two_case_variant_whose_payloads_overlap() {
        // Two cases, a 1-byte discriminant, payload i32 → align to 4, 4+4 = 8.
        let r = ValType::Result(
            Some(Box::new(ValType::I32)),
            Some(Box::new(ValType::I32)),
        );
        assert_eq!(alignment(&r), 4);
        assert_eq!(elem_size(&r), 8);
        // The cases OVERLAP — widening one does not add to the other.
        let wide = ValType::Result(
            Some(Box::new(ValType::I64)),
            Some(Box::new(ValType::I32)),
        );
        assert_eq!(elem_size(&wide), 16, "8-aligned discriminant + 8 payload");
    }

    /// `result<_, E>` — an ok arm carrying NOTHING, which is the shape every
    /// WASI 0.3.1 call returns and which this crate could not spell until the
    /// case payloads became optional.
    ///
    /// The old workaround declared the ok arm as some stand-in type, and this
    /// is exactly what that cost: a 1-byte error payload gives a 2-byte result,
    /// where a stand-in `i32` on the ok arm would have said 8.
    #[test]
    fn a_result_with_no_ok_payload_is_only_as_big_as_its_error() {
        let r = ValType::Result(None, Some(Box::new(ValType::Bool)));
        assert_eq!(alignment(&r), 1);
        assert_eq!(elem_size(&r), 2, "1-byte discriminant + 1-byte error");

        // And with nothing on either side it is the discriminant alone.
        let bare = ValType::Result(None, None);
        assert_eq!(elem_size(&bare), 1);
    }

    #[test]
    fn option_is_a_variant_with_an_empty_first_case() {
        // option<i32>: 1-byte discriminant, aligned to 4, + 4 = 8.
        let o = ValType::Option(Box::new(ValType::I32));
        assert_eq!(elem_size(&o), 8);
        // option<bool> collapses: 1-byte discriminant + 1-byte payload.
        assert_eq!(elem_size(&ValType::Option(Box::new(ValType::Bool))), 2);
    }

    #[test]
    fn align_to_rounds_up_and_leaves_aligned_values_alone() {
        assert_eq!(align_to(0, 4), 0);
        assert_eq!(align_to(1, 4), 4);
        assert_eq!(align_to(4, 4), 4);
        assert_eq!(align_to(5, 8), 8);
    }
}

/// `alignment_flags` / `elem_size_flags` — `CanonicalABI.md:2292` and `:2356`.
///
/// One packed integer: 1 byte up to 8 labels, 2 up to 16, 4 beyond. Both spec
/// functions are this same computation, which is why one helper serves both —
/// giving them separate bodies is how the two silently drift apart.
///
/// ⛔ The spec asserts `0 < n <= 32`. Beyond 32 the bits do not fit in the
/// `i32` the type flattens to, so a 33rd label would be silently dropped.
/// `lower_valspec` refuses that case; here the width simply saturates at 4.
pub fn flags_bytes(n: usize) -> u32 {
    if n <= 8 {
        1
    } else if n <= 16 {
        2
    } else {
        4
    }
}
