//! Result encoding for the CM canonical copy built-ins —
//! `CanonicalABI.md` §`canon stream.{read,write}` / §`canon future.{read,write}`.
//!
//! The spec implements all four built-ins as ONE function, `stream_copy` /
//! `future_copy`, parameterised by direction:
//!
//! ```python
//! def canon_stream_read(stream_t, opts, i, ptr, n):
//!   return stream_copy(ReadableStreamEnd, WritableBufferGuestImpl,
//!                      EventCode.STREAM_READ, stream_t, opts, i, ptr, n)
//! ```
//!
//! so the packing lives here rather than in each dispatch arm — four copies of
//! a bit-layout is four chances to disagree with a conforming runtime, and the
//! disagreement would be invisible to us because both ends would be ours.

/// `CopyResult` — `CanonicalABI.md`:
///
/// ```python
/// class CopyResult(IntEnum):
///   COMPLETED = 0
///   DROPPED = 1
///   CANCELLED = 2
/// ```
///
/// `DROPPED` means the *other* end is gone, so no further copies are possible.
/// `CANCELLED` is reachable only after this end issued a copy and then a
/// `cancel-{read,write}`; it tells wasm that ownership of the buffer has come
/// back. `COMPLETED` is everything else.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum CopyResult {
    Completed = 0,
    Dropped = 1,
    Cancelled = 2,
}

/// Which end of a `stream`/`future` a canonical built-in expects.
///
/// The spec's four `cancel-{read,write}` built-ins differ only in the
/// `ReadableStreamEnd` / `WritableStreamEnd` / `ReadableFutureEnd` /
/// `WritableFutureEnd` they check for, which is why `cancel_copy` takes the
/// end type as a parameter rather than existing four times.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EndKind {
    ReadableStream,
    WritableStream,
    ReadableFuture,
    WritableFuture,
}

impl EndKind {
    /// The canonical built-in name, for error messages that name the call the
    /// guest actually made.
    pub fn builtin_name(self) -> &'static str {
        match self {
            EndKind::ReadableStream => "stream.cancel-read",
            EndKind::WritableStream => "stream.cancel-write",
            EndKind::ReadableFuture => "future.cancel-read",
            EndKind::WritableFuture => "future.cancel-write",
        }
    }

    pub fn describe(self) -> &'static str {
        match self {
            EndKind::ReadableStream => "readable stream",
            EndKind::WritableStream => "writable stream",
            EndKind::ReadableFuture => "readable future",
            EndKind::WritableFuture => "writable future",
        }
    }
}

/// Returned instead of a packed result when the copy did not finish
/// synchronously. The real result arrives later as an `EventCode` event.
///
/// Only an ASYNC (`opts.async_`) copy can answer this — a synchronous one
/// suspends until an event exists, precisely so it can return a real payload.
pub const BLOCKED: u32 = 0xffff_ffff;

/// `Buffer.MAX_LENGTH` — the element count is packed into the top 28 bits, and
/// the spec fixes this ceiling *independently of the address type*: "even
/// though the number of elements copied is packed into an `addrtype`, the
/// maximum length of the buffer is fixed at `2^28 - 1`".
pub const MAX_LENGTH: u32 = (1 << 28) - 1;

/// `packed_result = result | (buffer.progress << 4)`.
///
/// ⚠ The count is NOT returned bare. A one-element successful read is `0x10`,
/// not `1` — low nibble is the `CopyResult`, everything above it is the
/// progress count. `documentation/wasi3vmchanges.md` described these as four
/// alternative flat values, which would have made every module we emit
/// disagree with every conforming runtime on the first element copied.
pub fn pack(result: CopyResult, progress: u32) -> u32 {
    debug_assert!(progress <= MAX_LENGTH, "progress exceeds Buffer.MAX_LENGTH");
    (result as u32) | (progress << 4)
}

/// Split a packed result back apart — for tests and for host code that
/// forwards a result it did not produce.
pub fn unpack(packed: u32) -> (u32, u32) {
    (packed & 0xf, packed >> 4)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn one_element_completed_is_0x10_not_1() {
        // The whole reason this module exists.
        assert_eq!(pack(CopyResult::Completed, 1), 0x10);
    }

    #[test]
    fn zero_progress_is_the_bare_code() {
        assert_eq!(pack(CopyResult::Completed, 0), 0);
        assert_eq!(pack(CopyResult::Dropped, 0), 1);
        assert_eq!(pack(CopyResult::Cancelled, 0), 2);
    }

    #[test]
    fn code_and_count_round_trip() {
        for &(code, n) in &[
            (CopyResult::Completed, 0u32),
            (CopyResult::Completed, 65536),
            (CopyResult::Dropped, 7),
            (CopyResult::Cancelled, MAX_LENGTH),
        ] {
            assert_eq!(unpack(pack(code, n)), (code as u32, n));
        }
    }

    #[test]
    fn blocked_is_not_a_valid_packed_result() {
        // 0xffff_ffff would unpack as code 0xf, which is not a CopyResult —
        // that is what lets one i32 carry both without ambiguity.
        let (code, _) = unpack(BLOCKED);
        assert!(code > CopyResult::Cancelled as u32);
    }
}
