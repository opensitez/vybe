//! Shared operand-stack slot initialization, independent of guest language.

/// Truncation retains normal reference destruction; callers still initialize
/// native Wasm locals and preserve their existing allocation/limit checks.
#[inline]
pub(crate) fn resize_undefined_values(values: &mut Vec<crate::value::Value>, len: usize) {
    resize_empty_values::<true>(values, len);
}

#[inline]
pub(crate) fn resize_null_values(values: &mut Vec<crate::value::Value>, len: usize) {
    resize_empty_values::<false>(values, len);
}

#[inline]
fn resize_empty_values<const UNDEFINED: bool>(values: &mut Vec<crate::value::Value>, len: usize) {
    if len <= values.len() {
        values.truncate(len);
    } else {
        let additional = len - values.len();
        values.extend(std::iter::repeat_with(|| if UNDEFINED {
            crate::value::Value::Undefined
        } else {
            crate::value::Value::Null
        }).take(additional));
    }
}

#[cfg(test)]
#[path = "stack_resize_tests.rs"]
mod stack_resize_tests;

#[cfg(test)]
#[path = "frame_invocation_tests.rs"]
mod frame_invocation_tests;

#[cfg(test)]
#[path = "value_cleanup_tests.rs"]
mod value_cleanup_tests;

#[cfg(test)]
#[path = "stack_growth_candidate_tests.rs"]
mod stack_growth_candidate_tests;

#[cfg(test)]
#[path = "numeric_roundtrip_frame_tests.rs"]
mod numeric_roundtrip_frame_tests;
