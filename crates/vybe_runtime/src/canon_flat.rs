//! `CanonicalABI.md` §Flattening — component types → core value types.
//!
//! The Canonical ABI would otherwise put every parameter and result in linear
//! memory. Flattening decomposes the ones that fit into core scalars so they
//! can travel in registers, and falls back to a single pointer when they do
//! not. This module is `flatten_type` / `flatten_types` / `flatten_functype`
//! and nothing else — the lifting and lowering that CONSUME a flattening are
//! separate (`canon_value` for the memory form, `canon_flat_values` for the
//! register form).
//!
//! ⚠The counts are the spec's, not ours to tune. `MAX_FLAT_RESULTS` is 1
//! "due to various parts of the toolchain (notably the C ABI) not yet being
//! able to express multi-value returns" — a temporary limit the spec expects to
//! lift, so it is named rather than folded into the code.

use crate::component::ValType;

/// `CanonicalABI.md`: `MAX_FLAT_PARAMS = 16`.
pub const MAX_FLAT_PARAMS: usize = 16;
/// `MAX_FLAT_ASYNC_PARAMS = 4` — an `async`-lowered function falls back to
/// memory sooner, because its arguments must survive across a suspension.
pub const MAX_FLAT_ASYNC_PARAMS: usize = 4;
/// `MAX_FLAT_RESULTS = 1`. See the module note: a toolchain limit, not a
/// design one.
pub const MAX_FLAT_RESULTS: usize = 1;

/// A core value type, the alphabet flattening maps into.
///
/// Deliberately NOT `ValType`: these are core wasm types (`i32`/`i64`/`f32`/
/// `f64`), and conflating them with component types is what makes a flattening
/// look like a lowering. `join` below only makes sense over this set.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CoreType {
    I32,
    I64,
    F32,
    F64,
}

/// Which side of the boundary a flattening is for — `flatten_functype`'s
/// `context` parameter. The two differ in where an oversized RESULT goes: a
/// lift returns a pointer, a lower takes one as an extra parameter.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FlattenContext {
    Lift,
    Lower,
}

/// The core signature a component functype flattens to.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CoreFuncType {
    pub params: Vec<CoreType>,
    pub results: Vec<CoreType>,
}

/// `join(a, b)` — the tightest core type that can carry either.
///
/// Variant cases each flatten differently, and rather than give up, the ABI
/// relies on the four core types being bit-castable and picks an
/// approximation. `i32`/`f32` join to `i32` (same width, reinterpreted);
/// anything else widens to `i64`.
pub fn join(a: CoreType, b: CoreType) -> CoreType {
    if a == b {
        return a;
    }
    match (a, b) {
        (CoreType::I32, CoreType::F32) | (CoreType::F32, CoreType::I32) => CoreType::I32,
        _ => CoreType::I64,
    }
}

/// `flatten_type(t, opts)`.
///
/// `ptr_type` is the address type of the `memory` option — `i32`, or `i64`
/// under memory64 (the spec's 🐘). It is a parameter rather than a constant
/// because a `string` flattens to TWO of them, and getting that wrong silently
/// halves or doubles a signature.
pub fn flatten_type(t: &ValType, ptr: CoreType) -> Vec<CoreType> {
    match t {
        ValType::Bool => vec![CoreType::I32],
        ValType::I32 => vec![CoreType::I32],
        ValType::I64 => vec![CoreType::I64],
        ValType::F64 => vec![CoreType::F64],
        // (ptr, length) — both of the memory's address type.
        ValType::String => vec![ptr, ptr],
        // An unfixed list is (ptr, length); this crate's `ValType::List` has no
        // fixed-length form, so there is no `flatten_list` length case to take.
        ValType::List(_) => vec![ptr, ptr],
        ValType::Record(fields) => {
            let mut flat = Vec::new();
            for (_, field) in fields {
                flat.extend(flatten_type(field, ptr));
            }
            flat
        }
        // `option` and `result` DESPECIALISE to variants — the spec flattens
        // them through the same path, so they must not get a shape of their
        // own here or a `result` would flatten differently from the two-case
        // variant it is defined to be.
        ValType::Option(inner) => flatten_variant_cases(
            &[None, Some(inner.as_ref().clone())],
            ptr,
        ),
        ValType::Result(ok, err) => flatten_variant_cases(
            &[ok.as_deref().cloned(), err.as_deref().cloned()],
            ptr,
        ),
        ValType::Variant(cases) => {
            let payloads: Vec<Option<ValType>> = cases.iter().map(|(_, t)| t.clone()).collect();
            flatten_variant_cases(&payloads, ptr)
        }
        // Handles and the async types are all one index.
        ValType::Own(_) | ValType::Borrow(_) => vec![CoreType::I32],
        ValType::Stream(_) | ValType::Future(_) => vec![CoreType::I32],
        // `Any` is not a component type and has no flattening, the same
        // position `canon_value` takes. Answering `[]` would silently drop a
        // parameter; answering `[i32]` would invent one.
        ValType::Any => Vec::new(),
    }
}

/// `flatten_variant(cases, opts)` — the discriminant, then the JOIN of every
/// case payload position-by-position.
///
/// Every case travels in the same static core types regardless of which one is
/// live, which is why `join` exists: position `i` must carry case A's `f32` and
/// case B's `i32` alike.
fn flatten_variant_cases(cases: &[Option<ValType>], ptr: CoreType) -> Vec<CoreType> {
    let mut flat: Vec<CoreType> = Vec::new();
    for case in cases {
        if let Some(t) = case {
            for (i, ft) in flatten_type(t, ptr).into_iter().enumerate() {
                if i < flat.len() {
                    flat[i] = join(flat[i], ft);
                } else {
                    flat.push(ft);
                }
            }
        }
    }
    let mut out = discriminant_core_type(cases.len());
    out.extend(flat);
    out
}

/// `flatten_type(discriminant_type(cases))` — the discriminant always flattens
/// to a single `i32`, whatever its stored width.
///
/// `canon_layout::discriminant_size` decides how many BYTES it occupies in
/// memory (1, 2 or 4); flattened, it is one core `i32` either way. Those two
/// facts are easy to conflate and answer different questions.
fn discriminant_core_type(_case_count: usize) -> Vec<CoreType> {
    vec![CoreType::I32]
}

/// `flatten_types(ts, opts)`.
pub fn flatten_types(ts: &[ValType], ptr: CoreType) -> Vec<CoreType> {
    ts.iter().flat_map(|t| flatten_type(t, ptr)).collect()
}

/// `flatten_functype(opts, ft, context)`.
///
/// `result` is the component function's single result type, `None` for a
/// function that returns nothing.
pub fn flatten_functype(
    params: &[ValType],
    result: Option<&ValType>,
    context: FlattenContext,
    is_async: bool,
    has_callback: bool,
    ptr: CoreType,
) -> CoreFuncType {
    let mut flat_params = flatten_types(params, ptr);
    let mut flat_results = match result {
        Some(t) => flatten_type(t, ptr),
        None => Vec::new(),
    };

    if !is_async {
        if flat_params.len() > MAX_FLAT_PARAMS {
            flat_params = vec![ptr];
        }
        if flat_results.len() > MAX_FLAT_RESULTS {
            match context {
                // A lift RETURNS the pointer…
                FlattenContext::Lift => flat_results = vec![ptr],
                // …a lower TAKES it, because the caller can often allocate the
                // space more cheaply (on its own stack) than `realloc` can.
                FlattenContext::Lower => {
                    flat_params.push(ptr);
                    flat_results = Vec::new();
                }
            }
        }
        return CoreFuncType {
            params: flat_params,
            results: flat_results,
        };
    }

    match context {
        FlattenContext::Lift => {
            if flat_params.len() > MAX_FLAT_PARAMS {
                flat_params = vec![ptr];
            }
            // With a callback the core function answers the packed event code;
            // without one it answers nothing and the result arrives later via
            // `task.return`.
            flat_results = if has_callback {
                vec![CoreType::I32]
            } else {
                Vec::new()
            };
        }
        FlattenContext::Lower => {
            // FOUR, not sixteen — an async lowering's arguments must outlive a
            // suspension, so the fallback to memory happens sooner.
            if flat_params.len() > MAX_FLAT_ASYNC_PARAMS {
                flat_params = vec![ptr];
            }
            if !flat_results.is_empty() {
                flat_params.push(ptr);
            }
            flat_results = vec![CoreType::I32];
        }
    }
    CoreFuncType {
        params: flat_params,
        results: flat_results,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn s(name: &str) -> String {
        name.to_string()
    }

    #[test]
    fn scalars_and_string_flatten_to_the_spec_shapes() {
        assert_eq!(flatten_type(&ValType::Bool, CoreType::I32), [CoreType::I32]);
        assert_eq!(flatten_type(&ValType::I64, CoreType::I32), [CoreType::I64]);
        assert_eq!(flatten_type(&ValType::F64, CoreType::I32), [CoreType::F64]);
        // A string is (ptr, length) — TWO, and both of the memory's address
        // type. One is the commonest way to get a signature quietly wrong.
        assert_eq!(
            flatten_type(&ValType::String, CoreType::I32),
            [CoreType::I32, CoreType::I32]
        );
        assert_eq!(
            flatten_type(&ValType::String, CoreType::I64),
            [CoreType::I64, CoreType::I64]
        );
    }

    #[test]
    fn a_record_flattens_its_fields_in_sequence() {
        let r = ValType::Record(vec![
            (s("a"), ValType::I32),
            (s("b"), ValType::String),
            (s("c"), ValType::F64),
        ]);
        assert_eq!(
            flatten_type(&r, CoreType::I32),
            [
                CoreType::I32,
                CoreType::I32,
                CoreType::I32,
                CoreType::F64
            ]
        );
    }

    /// The `join` rule is why a variant can be passed in fixed core types at
    /// all: position 1 must carry `f32` from one case and `i64` from another.
    #[test]
    fn variant_cases_join_position_by_position() {
        let v = ValType::Variant(vec![
            (s("a"), Some(ValType::I32)),
            (s("b"), Some(ValType::F64)),
        ]);
        // discriminant + join(i32, f64) = i64 (different widths widen)
        assert_eq!(
            flatten_type(&v, CoreType::I32),
            [CoreType::I32, CoreType::I64]
        );

        assert_eq!(join(CoreType::I32, CoreType::F32), CoreType::I32);
        assert_eq!(join(CoreType::F32, CoreType::I32), CoreType::I32);
        assert_eq!(join(CoreType::I32, CoreType::I32), CoreType::I32);
        assert_eq!(join(CoreType::I32, CoreType::F64), CoreType::I64);
    }

    /// `option` and `result` DESPECIALISE to variants — they must flatten
    /// identically to the two-case variant they are defined as, or a `result`
    /// crossing a boundary would disagree with itself.
    #[test]
    fn option_and_result_flatten_as_the_variants_they_despecialise_to() {
        let opt = ValType::Option(Box::new(ValType::I32));
        let equivalent = ValType::Variant(vec![(s("none"), None), (s("some"), Some(ValType::I32))]);
        assert_eq!(
            flatten_type(&opt, CoreType::I32),
            flatten_type(&equivalent, CoreType::I32)
        );

        // `result<_, error-code>`: only the error arm carries a payload, so the
        // flattening is discriminant + that payload.
        let res = ValType::Result(None, Some(Box::new(ValType::String)));
        assert_eq!(
            flatten_type(&res, CoreType::I32),
            [CoreType::I32, CoreType::I32, CoreType::I32]
        );
    }

    #[test]
    fn too_many_params_collapse_to_one_pointer() {
        let many: Vec<ValType> = (0..MAX_FLAT_PARAMS + 1).map(|_| ValType::I32).collect();
        let ft = flatten_functype(&many, None, FlattenContext::Lift, false, false, CoreType::I32);
        assert_eq!(ft.params, [CoreType::I32]);
        assert!(ft.results.is_empty());

        // Exactly at the limit is NOT over it.
        let exact: Vec<ValType> = (0..MAX_FLAT_PARAMS).map(|_| ValType::I32).collect();
        let ft = flatten_functype(&exact, None, FlattenContext::Lift, false, false, CoreType::I32);
        assert_eq!(ft.params.len(), MAX_FLAT_PARAMS);
    }

    /// The one asymmetry between the two contexts: an oversized result is
    /// RETURNED as a pointer by a lift and PASSED as one to a lower.
    #[test]
    fn an_oversized_result_returns_a_pointer_on_lift_and_takes_one_on_lower() {
        let big = ValType::Record(vec![(s("a"), ValType::I32), (s("b"), ValType::I32)]);

        let lift = flatten_functype(
            &[],
            Some(&big),
            FlattenContext::Lift,
            false,
            false,
            CoreType::I32,
        );
        assert!(lift.params.is_empty());
        assert_eq!(lift.results, [CoreType::I32]);

        let lower = flatten_functype(
            &[],
            Some(&big),
            FlattenContext::Lower,
            false,
            false,
            CoreType::I32,
        );
        assert_eq!(lower.params, [CoreType::I32]);
        assert!(lower.results.is_empty());
    }

    #[test]
    fn async_lower_falls_back_to_memory_at_four_params() {
        let five: Vec<ValType> = (0..MAX_FLAT_ASYNC_PARAMS + 1).map(|_| ValType::I32).collect();
        let ft = flatten_functype(&five, None, FlattenContext::Lower, true, false, CoreType::I32);
        assert_eq!(ft.params, [CoreType::I32]);
        // An async lowering always answers the packed i32.
        assert_eq!(ft.results, [CoreType::I32]);

        // A sync lowering of the same five is NOT collapsed — the async limit
        // is four, the sync limit sixteen.
        let ft = flatten_functype(&five, None, FlattenContext::Lower, false, false, CoreType::I32);
        assert_eq!(ft.params.len(), 5);
    }

    #[test]
    fn async_lift_answers_the_callback_code_only_with_a_callback() {
        let with = flatten_functype(&[], None, FlattenContext::Lift, true, true, CoreType::I32);
        assert_eq!(with.results, [CoreType::I32]);

        // Without a callback the result arrives via `task.return`, so the core
        // function answers nothing.
        let without = flatten_functype(&[], None, FlattenContext::Lift, true, false, CoreType::I32);
        assert!(without.results.is_empty());
    }
}
