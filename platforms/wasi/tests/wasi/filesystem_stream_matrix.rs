//! `read-via-stream` over the offset dimension, 0.3.1.
//!
//! Every case here goes through `stream_drain`, which builds ONE chunk doing
//! open-at → read-via-stream → `.at(0)` → `canon stream.read`. The
//! one-host-call-per-VM `invoke` helper the rest of these files use cannot
//! serve: a stream end is an index into that VM's handle table, so handing the
//! tuple out and calling back in yields a handle from a VM that is gone.

use std::path::PathBuf;
use std::sync::atomic::{AtomicU64, Ordering};

fn scratch_dir(label: &str) -> PathBuf {
    static COUNTER: AtomicU64 = AtomicU64::new(0);
    let id = COUNTER.fetch_add(1, Ordering::SeqCst);
    let dir = std::env::temp_dir().join(format!(
        "vybe-wasi-fs-stream-matrix-test-{}-{}-{}",
        std::process::id(),
        label,
        id
    ));
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).expect("scratch dir mkdir");
    dir
}

/// The offset dimension of `read-via-stream`, drained the 0.3.1 way.
///
/// The LENGTH dimension is gone, and its absence is the point: 0.2's
/// `input-stream.read(n)` took a byte count, while 0.3.1's stream carries
/// everything from `offset` and the guest's drain loop decides how much of it
/// to keep. These cases used to assert `(offset, length) -> slice`; there is
/// no length to pass any more, so each one asserts the whole suffix.
macro_rules! read_case {
    ($name:ident, $offset:expr, $expected:expr) => {
        #[test]
        fn $name() {
            let dir = scratch_dir(stringify!($name));
            std::fs::write(dir.join("payload.bin"), b"abcdef").unwrap();
            let bytes = crate::stream_drain::read_via_stream(&dir, "payload.bin", $offset);
            assert_eq!(bytes, $expected);
        }
    };
}

read_case!(read_from_offset_zero_returns_whole_file, 0.0, b"abcdef");
read_case!(read_from_offset_one_returns_suffix, 1.0, b"bcdef");
read_case!(read_from_offset_four_returns_remaining_suffix, 4.0, b"ef");

// Reading AT the end and PAST it are different conditions that happen to
// share an answer here: both drain empty. Kept as separate cases because a
// future short-read model has to keep them distinguishable.
read_case!(read_from_exact_end_returns_empty_array, 6.0, b"");
read_case!(read_from_beyond_end_returns_empty_array, 12.0, b"");

/// Each call gets its OWN stream, positioned by its own offset.
///
/// The 0.2 shape had one `input-stream` resource whose cursor advanced across
/// reads, so "sequential reads advance the position" was a property of the
/// resource. 0.3.1 has no such resource: `read-via-stream` behaves like
/// `pread`, so position is the argument and successive reads are independent.
/// Asserting the old property against the new function would assert a cursor
/// that no longer exists.
#[test]
fn each_read_is_positioned_independently() {
    let dir = scratch_dir("independent_positions");
    std::fs::write(dir.join("payload.bin"), b"abcdef").unwrap();
    assert_eq!(
        crate::stream_drain::read_via_stream(&dir, "payload.bin", 0.0),
        b"abcdef"
    );
    assert_eq!(
        crate::stream_drain::read_via_stream(&dir, "payload.bin", 3.0),
        b"def"
    );
    assert_eq!(
        crate::stream_drain::read_via_stream(&dir, "payload.bin", 0.0),
        b"abcdef",
        "a second read from 0 sees the whole file again — there is no cursor to have moved"
    );
}

/// `tuple<stream<u8>, future<result<_, error-code>>>`: element 1 resolves with
/// success on a clean read. `Null` is the `ok` case; an `error-code` arrives as
/// a string.
#[test]
fn clean_read_resolves_the_outcome_future_with_ok() {
    let dir = scratch_dir("outcome_ok");
    std::fs::write(dir.join("payload.bin"), b"abcdef").unwrap();
    let outcome = crate::stream_drain::read_via_stream_outcome(&dir, "payload.bin", 0.0);
    assert!(
        !matches!(outcome, vybe_runtime::value::Value::String(_)),
        "a clean read must not resolve the future with an error-code, got {outcome:?}"
    );
}
