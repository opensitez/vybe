//! Python `bytes` support.
//!
//! A `bytes` value is represented at runtime as a plain array of integers
//! (0–255), so `len`, indexing, iteration, concatenation (`+`) and repetition
//! (`*`) all reuse the existing array machinery for free. The only thing that
//! differs from a list is the *display*: `bytes` reprs as `b'…'` with
//! non-printable octets shown as `\xNN`.
//!
//! `b'…'` literals (and other statically-known-bytes expressions) are wrapped
//! by the walker in `__py_bytes__(<array>)` purely as a compile-time marker so
//! `expr_is_python_bytes` can find them; at runtime that marker is the identity
//! function below (the array flows straight through). Repr contexts are lowered
//! to a call to the `__vybe_bytes_repr` source prelude.

use vybe_bytecode::Chunk;

/// `__py_bytes__(array)` — identity marker. The single argument (the int array)
/// is already on the stack; leaving it there makes this a pure passthrough so
/// the value behaves exactly like the list it wraps for every non-display use.
pub fn emit_bytes_wrap(_chunks: &mut [Chunk], _current: usize, _argc: u8, _line: u32) {
    // Intentionally empty: the evaluated argument is the result.
}
