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
        // `CanonicalABI.md:3190` — every integer narrower than 64 bits
        // flattens to ONE core `i32`. The width is a MEMORY fact, not a flat
        // one: it decides how many bytes `load`/`store` move, and the range
        // check on the way in and out.
        ValType::S8 | ValType::U8 | ValType::S16 | ValType::U16 | ValType::I32 | ValType::U32 => {
            vec![CoreType::I32]
        }
        ValType::I64 | ValType::U64 => vec![CoreType::I64],
        // `f32` flattens to a core `f32`, NOT to `f64` — a demotion here would
        // round in a place the spec does not.
        ValType::F32 => vec![CoreType::F32],
        ValType::F64 => vec![CoreType::F64],
        // A `char` travels as its scalar value in one `i32`.
        ValType::Char => vec![CoreType::I32],
        // `CanonicalABI.md:3197` — flags flatten to ONE `i32` whatever their
        // memory width, because at most 32 labels fit and they are packed.
        ValType::Flags(_) => vec![CoreType::I32],
        // (ptr, length) — both of the memory's address type.
        ValType::String => vec![ptr, ptr],
        // An unfixed list is (ptr, length); this crate's `ValType::List` has no
        // fixed-length form, so there is no `flatten_list` length case to take.
        ValType::List(_) => vec![ptr, ptr],
        // 🔧 `flatten_type(t) * N` — N copies of the element's flat types,
        // NOT the (ptr, len) pair an unfixed list takes.
        ValType::ListFixed(elem, n) => {
            let one = flatten_type(elem, ptr);
            if one.is_empty() {
                return Vec::new();
            }
            let mut flat = Vec::with_capacity(one.len() * *n as usize);
            for _ in 0..*n {
                flat.extend(one.iter().copied());
            }
            flat
        }
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
        ValType::Option(inner) => flatten_variant_cases([None, Some(inner.as_ref())].into_iter(), ptr),
        ValType::Result(ok, err) => {
            flatten_variant_cases([ok.as_deref(), err.as_deref()].into_iter(), ptr)
        }
        ValType::Variant(cases) => {
            flatten_variant_cases(cases.iter().map(|(_, t)| t.as_ref()), ptr)
        }
        // Handles and the async types are all one index.
        // 📝 `error-context` flattens as its i32 handle, like own/borrow.
        ValType::Own(_) | ValType::Borrow(_) | ValType::ErrorContext => vec![CoreType::I32],
        ValType::Stream(_) | ValType::Future(_) => vec![CoreType::I32],
        // `Any` is not a component type and has no flattening, the same
        // position `canon_value` takes. Answering `[]` would silently drop a
        // parameter; answering `[i32]` would invent one.
        ValType::Any => Vec::new(),
    }
}

/// Count flattened values up to `limit + 1`, without expanding fixed lists.
/// The extra value distinguishes a signature exactly at the ABI limit from
/// one that must use linear memory.
pub fn flat_count_bounded(t: &ValType, limit: usize) -> usize {
    let cap = limit.saturating_add(1);
    match t {
        ValType::Any => 0,
        ValType::String | ValType::List(_) => 2.min(cap),
        ValType::ListFixed(elem, count) => flat_count_bounded(elem, limit)
            .saturating_mul(*count as usize)
            .min(cap),
        ValType::Record(fields) => fields.iter().fold(0usize, |total, (_, field)| {
            total.saturating_add(flat_count_bounded(field, limit)).min(cap)
        }),
        ValType::Option(inner) => 1usize
            .saturating_add(flat_count_bounded(inner, limit))
            .min(cap),
        ValType::Result(ok, err) => 1usize
            .saturating_add(
                ok.as_deref()
                    .into_iter()
                    .chain(err.as_deref())
                    .map(|case| flat_count_bounded(case, limit))
                    .max()
                    .unwrap_or(0),
            )
            .min(cap),
        ValType::Variant(cases) => 1usize
            .saturating_add(
                cases
                    .iter()
                    .filter_map(|(_, case)| case.as_ref())
                    .map(|case| flat_count_bounded(case, limit))
                    .max()
                    .unwrap_or(0),
            )
            .min(cap),
        _ => 1.min(cap),
    }
}

pub fn flat_types_count_bounded<'a>(
    types: impl IntoIterator<Item = &'a ValType>,
    limit: usize,
) -> usize {
    let cap = limit.saturating_add(1);
    types.into_iter().fold(0usize, |total, ty| {
        total.saturating_add(flat_count_bounded(ty, limit)).min(cap)
    })
}

/// `flatten_variant(cases, opts)` — the discriminant, then the JOIN of every
/// case payload position-by-position.
///
/// Every case travels in the same static core types regardless of which one is
/// live, which is why `join` exists: position `i` must carry case A's `f32` and
/// case B's `i32` alike.
pub(crate) fn flatten_variant_cases<'a>(
    cases: impl ExactSizeIterator<Item = Option<&'a ValType>>,
    ptr: CoreType,
) -> Vec<CoreType> {
    let mut flat: Vec<CoreType> = Vec::new();
    let case_count = cases.len();
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
    let mut out = discriminant_core_type(case_count);
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
    let param_limit = if is_async && context == FlattenContext::Lower {
        MAX_FLAT_ASYNC_PARAMS
    } else {
        MAX_FLAT_PARAMS
    };
    let mut flat_params = if flat_types_count_bounded(params, param_limit) > param_limit {
        vec![ptr]
    } else {
        flatten_types(params, ptr)
    };

    if !is_async {
        let flat_results = match result {
            Some(t) if flat_count_bounded(t, MAX_FLAT_RESULTS) > MAX_FLAT_RESULTS => {
                match context {
                    FlattenContext::Lift => vec![ptr],
                    FlattenContext::Lower => {
                        flat_params.push(ptr);
                        Vec::new()
                    }
                }
            }
            Some(t) => flatten_type(t, ptr),
            None => Vec::new(),
        };
        return CoreFuncType {
            params: flat_params,
            results: flat_results,
        };
    }

    match context {
        FlattenContext::Lift => {
            // With a callback the core function answers the packed event code;
            // without one it answers nothing and the result arrives later via
            // `task.return`.
            let flat_results = if has_callback {
                vec![CoreType::I32]
            } else {
                Vec::new()
            };
            CoreFuncType {
                params: flat_params,
                results: flat_results,
            }
        }
        FlattenContext::Lower => {
            // FOUR, not sixteen — an async lowering's arguments must outlive a
            // suspension, so the fallback to memory happens sooner.
            if result.is_some_and(|t| flat_count_bounded(t, 0) != 0) {
                flat_params.push(ptr);
            }
            CoreFuncType {
                params: flat_params,
                results: vec![CoreType::I32],
            }
        }
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
        assert_eq!(flatten_type(&ValType::U32, CoreType::I32), [CoreType::I32]);
        assert_eq!(flatten_type(&ValType::U64, CoreType::I32), [CoreType::I64]);
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
            [CoreType::I32, CoreType::I32, CoreType::I32, CoreType::F64]
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
        let ft = flatten_functype(
            &many,
            None,
            FlattenContext::Lift,
            false,
            false,
            CoreType::I32,
        );
        assert_eq!(ft.params, [CoreType::I32]);
        assert!(ft.results.is_empty());

        // Exactly at the limit is NOT over it.
        let exact: Vec<ValType> = (0..MAX_FLAT_PARAMS).map(|_| ValType::I32).collect();
        let ft = flatten_functype(
            &exact,
            None,
            FlattenContext::Lift,
            false,
            false,
            CoreType::I32,
        );
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
        let five: Vec<ValType> = (0..MAX_FLAT_ASYNC_PARAMS + 1)
            .map(|_| ValType::I32)
            .collect();
        let ft = flatten_functype(
            &five,
            None,
            FlattenContext::Lower,
            true,
            false,
            CoreType::I32,
        );
        assert_eq!(ft.params, [CoreType::I32]);
        // An async lowering always answers the packed i32.
        assert_eq!(ft.results, [CoreType::I32]);

        // A sync lowering of the same five is NOT collapsed — the async limit
        // is four, the sync limit sixteen.
        let ft = flatten_functype(
            &five,
            None,
            FlattenContext::Lower,
            false,
            false,
            CoreType::I32,
        );
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

    #[test]
    fn bounded_count_matches_flattening_without_expanding_large_fixed_lists() {
        let cases = [
            ValType::String,
            ValType::ListFixed(Box::new(ValType::I32), 0),
            ValType::ListFixed(Box::new(ValType::String), 3),
            ValType::Record(vec![(s("a"), ValType::I32), (s("b"), ValType::String)]),
            ValType::Option(Box::new(ValType::I64)),
            ValType::Result(Some(Box::new(ValType::I32)), Some(Box::new(ValType::String))),
            ValType::Variant(vec![(s("a"), None), (s("b"), Some(ValType::String))]),
        ];
        for ty in &cases {
            for limit in [0, 1, 4, 16] {
                assert_eq!(
                    flat_count_bounded(ty, limit),
                    flatten_type(ty, CoreType::I32).len().min(limit + 1),
                    "{ty:?} at limit {limit}"
                );
            }
        }
        let large = ValType::ListFixed(Box::new(ValType::U8), 268_435_455);
        assert_eq!(flat_count_bounded(&large, MAX_FLAT_RESULTS), 2);
        assert_eq!(flat_count_bounded(&large, MAX_FLAT_PARAMS), 17);
    }

    #[test]
    fn bounded_signature_selection_preserves_core_abi_shapes() {
        fn unbounded(
            params: &[ValType],
            result: Option<&ValType>,
            context: FlattenContext,
            is_async: bool,
            has_callback: bool,
            ptr: CoreType,
        ) -> CoreFuncType {
            let mut params = flatten_types(params, ptr);
            let mut results = result.map_or_else(Vec::new, |ty| flatten_type(ty, ptr));
            let limit = if is_async && context == FlattenContext::Lower {
                MAX_FLAT_ASYNC_PARAMS
            } else {
                MAX_FLAT_PARAMS
            };
            if params.len() > limit {
                params = vec![ptr];
            }
            if is_async {
                match context {
                    FlattenContext::Lift => {
                        results = if has_callback { vec![CoreType::I32] } else { vec![] };
                    }
                    FlattenContext::Lower => {
                        if !results.is_empty() {
                            params.push(ptr);
                        }
                        results = vec![CoreType::I32];
                    }
                }
            } else if results.len() > MAX_FLAT_RESULTS {
                match context {
                    FlattenContext::Lift => results = vec![ptr],
                    FlattenContext::Lower => {
                        params.push(ptr);
                        results.clear();
                    }
                }
            }
            CoreFuncType { params, results }
        }

        let shapes = vec![
            ValType::Any,
            ValType::I32,
            ValType::String,
            ValType::ListFixed(Box::new(ValType::I64), 0),
            ValType::ListFixed(Box::new(ValType::I64), 4),
            ValType::ListFixed(Box::new(ValType::I64), 16),
            ValType::ListFixed(Box::new(ValType::I64), 17),
            ValType::Result(Some(Box::new(ValType::F32)), Some(Box::new(ValType::I64))),
        ];
        for params in &shapes {
            for result in &shapes {
                for context in [FlattenContext::Lift, FlattenContext::Lower] {
                    for is_async in [false, true] {
                        for has_callback in [false, true] {
                            for ptr in [CoreType::I32, CoreType::I64] {
                                let params = std::slice::from_ref(params);
                                let actual = flatten_functype(
                                    params, Some(result), context, is_async, has_callback, ptr,
                                );
                                let expected = unbounded(
                                    params, Some(result), context, is_async, has_callback, ptr,
                                );
                                assert_eq!(actual, expected, "{params:?} {result:?} {context:?} async={is_async} callback={has_callback} ptr={ptr:?}");
                            }
                        }
                    }
                }
            }
        }

        let huge = ValType::ListFixed(Box::new(ValType::I64), 1_000_000);
        for context in [FlattenContext::Lift, FlattenContext::Lower] {
            let signature = flatten_functype(&[huge.clone()], Some(&huge), context, false, false, CoreType::I64);
            assert_eq!(signature.params, if context == FlattenContext::Lower { vec![CoreType::I64, CoreType::I64] } else { vec![CoreType::I64] });
            assert_eq!(signature.results, if context == FlattenContext::Lift { vec![CoreType::I64] } else { vec![] });
        }
        let zero_width = ValType::ListFixed(Box::new(ValType::Record(vec![])), 268_435_455);
        let signature = flatten_functype(
            &[zero_width], None, FlattenContext::Lift, false, false, CoreType::I32,
        );
        assert!(signature.params.is_empty());
    }
}
