//! I/O compilation — WASI-compatible print, input, file operations.
//!
//! Print uses `web:console.log` (WHATWG Console Standard — `log(...data)`
//! is VARIADIC BY SPEC; each datum rendered, space-joined). The strict
//! `wasi:logging/logging.log(level, context, message)` remains for code
//! calling the WASI interface explicitly.
//! Input uses `wasi:cli/stdin.read-via-stream` drained with `canon stream.read`
//! — `wasi:cli@0.3.1`'s only stdin function. The 0.2 pair this header used to
//! name (`get-stdin` + `[method]input-stream.blocking-read`) is gone: `stdio.wit`
//! in the 0.3.1 tree declares `read-via-stream` and nothing else, and there is
//! no `io` proposal under `proposals/WASI/proposals/` at all.
//! File I/O uses `wasi:filesystem/*` imports.
//!
//! Output BUFFERING lives here too, for one structural reason: this module owns
//! the write. A buffer that some writers respect and others bypass is not a
//! buffer, and that is exactly the bug that existed while each language kept its
//! own — PHP's `echo` checked the buffer and its `var_dump` did not. Routing
//! every write through [`emit_write_or_buffer`] makes capture correct by
//! construction for every writer in every language, present and future.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// Emit print/log. Stack: [arg1, ..., argN] → []
/// Routes to `web:console.log` — WHATWG `log(...data)`, variadic by spec
/// (the host renders each datum and joins with a single space).
pub fn emit_print(chunk: &mut Chunk, arg_count: u8, line: u32) {
    let idx = chunk.add_import("web:console", "log");
    chunk.emit_call(idx, arg_count, line);
}

/// Emit print using a pre-resolved import index.
pub fn emit_print_with_import(chunk: &mut Chunk, import_idx: u16, arg_count: u8, line: u32) {
    chunk.emit_call(import_idx, arg_count, line);
}

/// Emit a raw byte write to stdout — NO implicit newline, unlike
/// `wasi:logging/logging.log` which is one line-oriented record per call.
///
/// Composes the WASI 0.3 surface: `canon stream.new` (a "canon"-module
/// import, called via spec `call`) creates a `stream<u8>` as (readable,
/// writable) i32 handles, the contents go in via `canon stream.write`,
/// the writable end closes (EOF), and the readable end is passed to
/// `wasi:cli/stdout.write-via-stream(data: stream<u8>)`.
/// The returned `future<result<_, error-code>>` is discarded; both handle
/// table entries are dropped afterwards per the canonical ABI.
///
/// Stack: [] → []. `push_contents` emits the string to write while the
/// writable handle is on the stack. `rd_slot`/`wr_slot` are caller-defined
/// scratch locals for the two canon handles; `write_idx` is the resolved
/// `wasi:cli/stdout.write-via-stream` import.
pub fn emit_write_stdout_with_imports(
    chunk: &mut Chunk,
    write_idx: u16,
    rd_slot: u16,
    wr_slot: u16,
    line: u32,
    push_contents: impl FnOnce(&mut Chunk),
) {
    // The `wasi:cli` shape: `write-via-stream(data)`, one argument, and the
    // future it answers is discarded.
    emit_write_via_stream(chunk, rd_slot, wr_slot, line, push_contents, |chunk, rd| {
        chunk.emit_op_u16(Op::LOCAL_GET, rd, line);
        chunk.emit_call(write_idx, 1, line);
        chunk.emit_op(Op::DROP, line);
    });
}

/// The transport under every `write-via-stream`, whatever its signature.
///
/// `emit_sink_call` receives the READABLE end's slot and emits the actual
/// import call. That is the only thing the sinks disagree about:
///
///     wasi:cli/{stdout,stderr}.write-via-stream(data)              → future
///     wasi:filesystem/types.[method]descriptor
///         .write-via-stream(this, data, offset)                    → future
///
/// Same "quirks above, ONE transport below" split the print path already uses:
/// minting the stream, marshalling the bytes into linear memory, writing and
/// dropping both ends is identical for all of them, so it lives here once.
pub fn emit_write_via_stream(
    chunk: &mut Chunk,
    rd_slot: u16,
    wr_slot: u16,
    line: u32,
    push_contents: impl FnOnce(&mut Chunk),
    emit_sink_call: impl FnOnce(&mut Chunk, u16),
) {
    emit_payload_via_stream(
        chunk,
        rd_slot,
        wr_slot,
        line,
        Payload::Text,
        push_contents,
        emit_sink_call,
    );
}

/// What `push_contents` leaves on the stack, and therefore how it is marshalled
/// into linear memory before `canon stream.write`.
///
/// The distinction is not cosmetic. A `Payload::Bytes` array sent down the
/// `Text` path is `TextEncoder`'d — the file receives the array's DECIMAL
/// RENDERING, `"72,101,108"`, with a plausible length and a successful return.
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum Payload {
    /// A string. Encoded UTF-8.
    Text,
    /// An array of byte-valued numbers. Copied verbatim.
    Bytes,
}

/// [`emit_write_via_stream`] with the payload kind spelled out.
pub fn emit_payload_via_stream(
    chunk: &mut Chunk,
    rd_slot: u16,
    wr_slot: u16,
    line: u32,
    payload: Payload,
    push_contents: impl FnOnce(&mut Chunk),
    emit_sink_call: impl FnOnce(&mut Chunk, u16),
) {
    // CM canonical built-ins are (core func) IMPORTS under module "canon"
    // (the 0xF0 instruction prefix is retired) — spec `call` throughout.
    let stream_new = chunk.add_import("canon", "stream.new");
    let stream_write = chunk.add_import("canon", "stream.write");
    let drop_wr = chunk.add_import("canon", "stream.drop-writable");
    let drop_rd = chunk.add_import("canon", "stream.drop-readable");
    // canon stream.new → ONE i64: `ri | (wi << 32)`, readable in the low 32
    // bits and writable in the high 32 (`CanonicalABI.md` §canon
    // {stream,future}.new — `$f` is given type `(func (result i64))`).
    // This used to take two stack values, which no conforming module could
    // have produced.
    // ⛔ THE PACKED HANDLE NEEDS ITS OWN SLOT. This parked the i64 in
    // `rd_slot` and then OVERWROTE that same slot with the 32-bit readable
    // end — one slot holding an i64 and later an i32. Our VM does not care;
    // WASM does, because a local has ONE declared type, and `rd_slot` is an
    // `externref` like every other. V8: `i64.shr_u[0] expected type i64, found
    // local.get of type externref`. It is also what made this look like it
    // needed general typed-local inference: no per-slot type could ever have
    // described a slot that changes type halfway through.
    let packed = chunk.i64_scratch();
    chunk.emit_call(stream_new, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, packed, line);
    // writable = high 32
    chunk.emit_op_u16(Op::LOCAL_GET, packed, line);
    chunk.emit_i64_const(32, line);
    chunk.emit_op(Op::I64_SHR_U, line);
    chunk.emit_op(Op::I32_WRAP_I64, line);
    chunk.emit_op_u16(Op::LOCAL_SET, wr_slot, line);
    // readable = low 32
    chunk.emit_op_u16(Op::LOCAL_GET, packed, line);
    chunk.emit_op(Op::I32_WRAP_I64, line);
    chunk.emit_op_u16(Op::LOCAL_SET, rd_slot, line);
    // canon stream.write(handle, ptr, n) — `CanonicalABI.md` §canon
    // stream.{read,write}: elements come FROM LINEAR MEMORY, they are not
    // handed over as a value. So the contents are marshalled first —
    // `__vybe_canon_store_utf8` encodes, allocates and copies, answering
    // `ptr | (len << 32)` in one i64.
    //
    // This used to push the string itself and call with argc 2, which is not a
    // signature any conforming component could satisfy. It worked only because
    // both ends were ours.
    // Marshal the contents into linear memory, INLINE. `emit_store_utf8`
    // splices the encode/allocate/copy at this call site and answers the two
    // scratch slots — no `__stdlib_*` helper, no global, no link-time
    // ordering.
    let (ptr, n) = match payload {
        Payload::Text => {
            crate::primitives::canon_marshal::emit_store_utf8(chunk, line, push_contents)
        }
        Payload::Bytes => {
            crate::primitives::canon_marshal::emit_store_byte_array(chunk, line, push_contents)
        }
    };
    chunk.emit_op_u16(Op::LOCAL_GET, wr_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ptr, line);
    chunk.emit_op_u16(Op::LOCAL_GET, n, line);
    chunk.emit_call(stream_write, 3, line);
    // The packed CopyResult is not inspected here: a write into a stream whose
    // reader is the host cannot partially fail, and stdout has nowhere to
    // report it. A guest that cares reads the result.
    chunk.emit_op(Op::DROP, line);
    // stream.drop-writable(wr) — signals EOF. MUST precede the sink call: the
    // sink drains until the writable end is gone, so handing it a still-open
    // stream is a wait for bytes that will never come.
    chunk.emit_op_u16(Op::LOCAL_GET, wr_slot, line);
    chunk.emit_call(drop_wr, 1, line);
    emit_sink_call(chunk, rd_slot);
    // stream.drop-readable(rd)
    chunk.emit_op_u16(Op::LOCAL_GET, rd_slot, line);
    chunk.emit_call(drop_rd, 1, line);
}

/// Drain a Component Model `stream<u8>` to a BYTE ARRAY. Stack: [handle] → [array].
///
/// THE read, and the mirror of [`emit_write_stdout_with_imports`]. WASI 0.3
/// deleted `wasi:io`, so `[method]input-stream.blocking-read` — which answered
/// a `list<u8>` as a VALUE — has no replacement in the WASI namespace at all.
/// A stream is a Component Model type now, and the only conforming way to read
/// one is `canon stream.read`.
///
/// `canon stream.read(handle, ptr, n)` copies elements INTO LINEAR MEMORY and
/// answers a packed `CopyResult`, so the bytes have to be fetched back out —
/// that is what `canon_marshal::emit_load_utf8` is for.
///
/// The packed result is `result | (progress << 4)` (`CanonicalABI.md` §canon
/// stream.{read,write}), so a one-byte read is `0x10`, NOT `1`. Two outcomes
/// reach this loop:
///   - `COMPLETED` (0) with a count — append and go round again.
///   - `DROPPED` (1) — the writer is gone; this is clean EOF.
///
/// `BLOCKED` is NOT among them, and this loop used to break on it. Only the
/// `async` variant of `canon stream.read` may answer BLOCKED; the synchronous
/// one — the only one this emits — SUSPENDS until the copy can proceed
/// (`CanonicalABI.md` §canon stream.{read,write}, and the `StreamRead` arm in
/// `dispatch.rs`). Breaking was defensible while the runtime answered BLOCKED
/// from the sync form: a spin-retry would have trapped, because a blocked read
/// leaves the end `Copying` and `stream.read` traps unless it is `Idle`. But
/// the cost of breaking is a SHORT READ reported as a complete one. On a file,
/// whose bytes are already buffered, that never showed. On a socket, where
/// nothing-ready-yet is the ordinary state, it is silent truncation.
///
/// Answers BYTES rather than text because the encoding is the CALLER's
/// business — an HTTP body is latin-1 one-char-per-byte so binary uploads
/// round-trip, stdout/stdin are UTF-8. Same transport, decode on top:
/// [`emit_read_stream_to_string`] is this plus `emit_decode_utf8`.
pub fn emit_read_stream_to_bytes(chunk: &mut Chunk, line: u32) {
    /// One read's buffer. Deliberately NOT the old `blocking-read`'s
    /// single-shot 64MB: that request went through the shared bump allocator,
    /// which would grow linear memory by 1024 pages at once and never free it.
    /// A bounded buffer plus the loop reads any body of any size.
    const CHUNK_BYTES: i32 = 65536;

    let stream_read = chunk.add_import("canon", "stream.read");
    let drop_rd = chunk.add_import("canon", "stream.drop-readable");

    let handle = chunk.alloc_scratch(1);
    let out = chunk.alloc_scratch(1);
    let packed = chunk.alloc_scratch(1);
    let count = chunk.alloc_scratch(1);
    let n = chunk.alloc_scratch(1);

    chunk.emit_op_u16(Op::LOCAL_SET, handle, line);
    crate::primitives::canon_marshal::emit_new_bytes(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, out, line);

    chunk.emit_i32_const(CHUNK_BYTES, line);
    chunk.emit_op_u16(Op::LOCAL_SET, n, line);
    let ptr = crate::primitives::canon_marshal::emit_alloc(chunk, line, n);

    let done = chunk.emit_block(line);
    let (drain, _) = chunk.emit_loop_s(line);

    chunk.emit_op_u16(Op::LOCAL_GET, handle, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ptr, line);
    chunk.emit_i32_const(CHUNK_BYTES, line);
    chunk.emit_call(stream_read, 3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, packed, line);

    // count = packed >> 4 (unsigned: the count occupies the top 28 bits).
    chunk.emit_op_u16(Op::LOCAL_GET, packed, line);
    chunk.emit_i32_const(4, line);
    chunk.emit_op(Op::I32_SHR_U, line);
    chunk.emit_op_u16(Op::LOCAL_SET, count, line);

    chunk.emit_op_u16(Op::LOCAL_GET, count, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_GT_S, line);
    chunk.emit_if(line);
    crate::primitives::canon_marshal::emit_append_bytes(chunk, line, out, ptr, count);
    chunk.emit_end(line);

    // Anything but COMPLETED (low nibble 0) ends the drain: DROPPED is EOF and
    // CANCELLED cannot occur here (no cancel was issued).
    chunk.emit_op_u16(Op::LOCAL_GET, packed, line);
    chunk.emit_i32_const(0xf, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_br_if(1, line);

    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(drain);
    chunk.emit_end(line);
    chunk.patch_block(done);

    chunk.emit_op_u16(Op::LOCAL_GET, handle, line);
    chunk.emit_call(drop_rd, 1, line);

    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
}

/// ONE `canon stream.read` of at most `max` bytes. Stack: [handle, max] → [bytes].
///
/// The bounded read, and NOT [`emit_read_stream_to_bytes`]. The difference is
/// the whole difference between a file and a socket.
///
/// A drain loops until the writer is gone. That terminates on a file, whose
/// bytes are all already there, and on a stream a host closed after filling
/// it. It does NOT terminate on a live socket: `receive`'s stream is left
/// OPEN with a producer precisely so that "the peer has not sent anything yet"
/// stays distinct from "the peer is gone", so a drain reads what has arrived,
/// goes round again, and SUSPENDS until the connection ends. `recv(2)` returns
/// as soon as anything is available — a drain would answer the whole
/// conversation, once, at disconnect.
///
/// So this reads exactly once and answers what that read produced. An empty
/// array means end of stream (`DROPPED`); the suspend inside `canon
/// stream.read` is what makes a blocking `recv` block, and the reason no
/// `pollable` is needed to spell the same wait.
///
/// The handle is deliberately NOT dropped: a socket is read again.
pub fn emit_read_stream_chunk(chunk: &mut Chunk, line: u32) {
    let stream_read = chunk.add_import("canon", "stream.read");

    // Args arrive in call order, so `max` is on top.
    let n = chunk.alloc_scratch(1);
    let handle = chunk.alloc_scratch(1);
    let out = chunk.alloc_scratch(1);
    let packed = chunk.alloc_scratch(1);
    let count = chunk.alloc_scratch(1);

    chunk.emit_op_u16(Op::LOCAL_SET, n, line);
    chunk.emit_op_u16(Op::LOCAL_SET, handle, line);

    crate::primitives::canon_marshal::emit_new_bytes(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, out, line);

    let ptr = crate::primitives::canon_marshal::emit_alloc(chunk, line, n);

    chunk.emit_op_u16(Op::LOCAL_GET, handle, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ptr, line);
    chunk.emit_op_u16(Op::LOCAL_GET, n, line);
    chunk.emit_call(stream_read, 3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, packed, line);

    // count = packed >> 4 — `result | (progress << 4)`, so one byte is 0x10.
    chunk.emit_op_u16(Op::LOCAL_GET, packed, line);
    chunk.emit_i32_const(4, line);
    chunk.emit_op(Op::I32_SHR_U, line);
    chunk.emit_op_u16(Op::LOCAL_SET, count, line);

    chunk.emit_op_u16(Op::LOCAL_GET, count, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_GT_S, line);
    chunk.emit_if(line);
    crate::primitives::canon_marshal::emit_append_bytes(chunk, line, out, ptr, count);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
}

/// Read ONE `own<T>` element out of a `stream<own<T>>`.
/// Stack: [handle] → [representation | null].
///
/// This is what `accept` became. 0.3.1 deleted `wasi:sockets/tcp.accept`
/// outright: `listen` answers `stream<tcp-socket>` and "a connection arrived"
/// is an ELEMENT on that stream, so accepting one is a read, not a call. The
/// wait is the read's own suspend, which is why no `pollable` is needed to
/// spell it either.
///
/// A resource is 4 bytes on the wire — `own<T>`/`borrow<T>` lower as an i32
/// index into the handle table (`canon_layout::elem_size`). So the element
/// copied into linear memory is a HANDLE, and `canon resource.rep` turns it
/// back into the representation its owner chose. Per §`canon resource.rep`
/// that representation is an i32 and nothing else — a host cannot smuggle an
/// object through it, which is why the thing on the other side has to be
/// resolvable FROM an i32.
///
/// Null means end of stream. For a listener that is a fatal error rather than
/// the ordinary case: §`listen` calls its result "a single perpetual stream
/// that should only close on fatal errors".
pub fn emit_read_stream_handle(chunk: &mut Chunk, line: u32) {
    /// One `own<T>`, in bytes.
    const HANDLE_BYTES: i32 = 4;

    let stream_read = chunk.add_import("canon", "stream.read");
    let resource_rep = chunk.add_import("canon", "resource.rep");

    let handle = chunk.alloc_scratch(1);
    let packed = chunk.alloc_scratch(1);
    let n = chunk.alloc_scratch(1);

    chunk.emit_op_u16(Op::LOCAL_SET, handle, line);
    chunk.emit_i32_const(HANDLE_BYTES, line);
    chunk.emit_op_u16(Op::LOCAL_SET, n, line);
    // `emit_alloc` answers an 8-aligned address, which covers a handle's
    // 4-byte alignment — `canon stream.read` TRAPS on a misaligned element
    // buffer rather than writing it crooked.
    let ptr = crate::primitives::canon_marshal::emit_alloc(chunk, line, n);

    chunk.emit_op_u16(Op::LOCAL_GET, handle, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ptr, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_call(stream_read, 3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, packed, line);

    // `result | (progress << 4)`, so ONE element reads back as 0x10, not 1.
    chunk.emit_op_u16(Op::LOCAL_GET, packed, line);
    chunk.emit_i32_const(4, line);
    chunk.emit_op(Op::I32_SHR_U, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_GT_S, line);
    chunk.emit_if_value(line);
    {
        chunk.emit_op_u16(Op::LOCAL_GET, ptr, line);
        chunk.emit_op(Op::I32_LOAD, line);
        chunk.emit_call(resource_rep, 1, line);
    }
    chunk.emit_else(line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunk.emit_end(line);
}

/// Read ONE `own<T>` element WITHOUT suspending. Stack: [handle] → [rep | null].
///
/// The non-blocking twin of [`emit_read_stream_handle`], and what `O_NONBLOCK`
/// actually means: `async` is canonopt 0x06 on the `canon stream.read` row, so
/// the ASYNC form may answer `BLOCKED` where the synchronous one must suspend.
/// `stream.read@0` names the async row by CANONIDX — the documented fallback
/// for a front end with no binder concept. Nothing here is Component-Model
/// specific: `compile_with_imports` lowers a canon section for every language.
///
/// ⚠ A BLOCKED read leaves the end COPYING — the copy IS in flight — and only
/// `stream.cancel-read` returns it to IDLE. Skipping that makes the NEXT read
/// on the same stream trap with "not IDLE", so the cancel is not optional
/// tidying: EAGAIN is two calls, not one.
pub fn emit_try_read_stream_handle(chunk: &mut Chunk, line: u32) {
    const HANDLE_BYTES: i32 = 4;
    const BLOCKED: i32 = -1;

    let stream_read = chunk.add_import("canon", "stream.read@0");
    let cancel_read = chunk.add_import("canon", "stream.cancel-read");
    let resource_rep = chunk.add_import("canon", "resource.rep");

    let handle = chunk.alloc_scratch(1);
    let packed = chunk.alloc_scratch(1);
    let n = chunk.alloc_scratch(1);

    chunk.emit_op_u16(Op::LOCAL_SET, handle, line);
    chunk.emit_i32_const(HANDLE_BYTES, line);
    chunk.emit_op_u16(Op::LOCAL_SET, n, line);
    let ptr = crate::primitives::canon_marshal::emit_alloc(chunk, line, n);

    chunk.emit_op_u16(Op::LOCAL_GET, handle, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ptr, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_call(stream_read, 3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, packed, line);

    // BLOCKED (0xffff_ffff) — nothing was queued. Reclaim the buffer so the
    // end is IDLE again, then answer "no element".
    chunk.emit_op_u16(Op::LOCAL_GET, packed, line);
    chunk.emit_i32_const(BLOCKED, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_if_value(line);
    {
        chunk.emit_op_u16(Op::LOCAL_GET, handle, line);
        chunk.emit_call(cancel_read, 1, line);
        chunk.emit_op(Op::DROP, line);
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    }
    chunk.emit_else(line);
    {
        // `result | (progress << 4)` — one element reads back as 0x10.
        chunk.emit_op_u16(Op::LOCAL_GET, packed, line);
        chunk.emit_i32_const(4, line);
        chunk.emit_op(Op::I32_SHR_U, line);
        chunk.emit_i32_const(0, line);
        chunk.emit_op(Op::I32_GT_S, line);
        chunk.emit_if_value(line);
        chunk.emit_op_u16(Op::LOCAL_GET, ptr, line);
        chunk.emit_op(Op::I32_LOAD, line);
        chunk.emit_call(resource_rep, 1, line);
        chunk.emit_else(line);
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        chunk.emit_end(line);
    }
    chunk.emit_end(line);
}

/// Hand a byte array over AS a Component Model `stream<u8>`.
/// Stack: [bytes] → [readable handle].
///
/// The mirror of [`emit_read_stream_to_bytes`], and the other half of what
/// replaced `wasi:io`. 0.3.1 deleted `[method]output-stream.blocking-write-and-
/// flush`, which took a `list<u8>` as a VALUE; every 0.3.1 sink that used to
/// accept bytes now takes a `stream<u8>` PARAMETER instead —
/// `wasi:sockets/types.[method]tcp-socket.send(data: stream<u8>)`,
/// `wasi:cli/stdout.write-via-stream(data: stream<u8>)`. So a caller does not
/// write bytes to a sink any more; it produces a stream and passes it.
///
/// That is all this is: mint the pair, marshal the bytes into linear memory,
/// `canon stream.write` them into the writable end, and drop that end so the
/// reader sees a clean EOF rather than waiting for a writer that is done. What
/// is left on the stack is the READABLE end — an ordinary argument, which is
/// why a sink taking one needs no special emitter and stays a plain profile
/// row.
///
/// Distinct from [`emit_payload_via_stream`], which does the same transport but
/// CALLS the sink itself and is therefore tied to one. Splitting the value out
/// is what makes this reusable by any `stream<u8>` parameter in any interface.
pub fn emit_bytes_to_stream(chunk: &mut Chunk, line: u32) {
    let stream_new = chunk.add_import("canon", "stream.new");
    let stream_write = chunk.add_import("canon", "stream.write");
    let drop_wr = chunk.add_import("canon", "stream.drop-writable");

    let src = chunk.alloc_scratch(1);
    let rd = chunk.alloc_scratch(1);
    let wr = chunk.alloc_scratch(1);

    chunk.emit_op_u16(Op::LOCAL_SET, src, line);

    // canon stream.new → ONE i64: `ri | (wi << 32)` (`CanonicalABI.md` §canon
    // {stream,future}.new). Same unpacking as `emit_payload_via_stream`.
    // Own slot for the packed i64 — see `emit_payload_via_stream`.
    let packed = chunk.i64_scratch();
    chunk.emit_call(stream_new, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, packed, line);
    chunk.emit_op_u16(Op::LOCAL_GET, packed, line);
    chunk.emit_i64_const(32, line);
    chunk.emit_op(Op::I64_SHR_U, line);
    chunk.emit_op(Op::I32_WRAP_I64, line);
    chunk.emit_op_u16(Op::LOCAL_SET, wr, line);
    chunk.emit_op_u16(Op::LOCAL_GET, packed, line);
    chunk.emit_op(Op::I32_WRAP_I64, line);
    chunk.emit_op_u16(Op::LOCAL_SET, rd, line);

    // `Payload::Bytes`, never `Text`: these bytes are a socket payload, and the
    // `Text` path would send the array's DECIMAL RENDERING with a plausible
    // length and a successful return.
    let (ptr, n) = crate::primitives::canon_marshal::emit_store_byte_array(chunk, line, |chunk| {
        chunk.emit_op_u16(Op::LOCAL_GET, src, line);
    });

    chunk.emit_op_u16(Op::LOCAL_GET, wr, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ptr, line);
    chunk.emit_op_u16(Op::LOCAL_GET, n, line);
    chunk.emit_call(stream_write, 3, line);
    // The packed CopyResult is dropped for the same reason the sink form drops
    // it: the whole payload is already in linear memory and the reader is the
    // host, so there is no partial-write outcome a caller could act on.
    chunk.emit_op(Op::DROP, line);

    // Drop-writable BEFORE the handle is used. The reader drains until the
    // writable end is gone; handing over a still-open stream is a wait for
    // bytes that will never come.
    chunk.emit_op_u16(Op::LOCAL_GET, wr, line);
    chunk.emit_call(drop_wr, 1, line);

    chunk.emit_op_u16(Op::LOCAL_GET, rd, line);
}

/// Drain a Component Model `stream<u8>` and decode it as UTF-8.
/// Stack: [handle] → [string].
///
/// The text-shaped read: [`emit_read_stream_to_bytes`] for the transport, one
/// `TextDecoder` pass over the whole run for the encoding. Callers that need a
/// different encoding (an HTTP body wants latin-1) call the bytes form and
/// decode themselves rather than getting a second drain loop.
pub fn emit_read_stream_to_string(chunk: &mut Chunk, line: u32) {
    emit_read_stream_to_bytes(chunk, line);
    crate::primitives::canon_marshal::emit_decode_utf8(chunk, line);
}

/// Emit print to stderr. Stack: [message] → []
/// WHATWG `console.error(...data)` — the stderr stream of the same
/// console surface.
pub fn emit_print_error(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("web:console", "error");
    chunk.emit_call(idx, 1, line);
}

/// Emit readline — ONE line from stdin, newline stripped. Stack: [] → [string]
///
/// THE stdin read: python `input()`, C# `Console.ReadLine()`, libc's
/// `intrinsic:readline` and every other line-at-a-time surface land here.
///
/// `wasi:cli/stdin.read-via-stream() -> tuple<stream<u8>, future<result<_,
/// error-code>>>`, drained with `canon stream.read`. It was `get-stdin` +
/// `wasi:io/streams.[method]input-stream.blocking-read`, which asked a package
/// 0.3 DELETED for a `list<u8>` — and in this host answered a whole STRING,
/// already split into a line, so the line semantics lived in the host where no
/// language could see them.
///
/// **The line buffer is the reason this is not a one-liner.** A stream is
/// bytes, not lines, and how many arrive per read is the transport's business:
/// piped stdin hands over the entire file at once, a terminal hands over what
/// was typed. So the surplus past the first newline is carried in a global and
/// consumed by later calls. Without it the first `input()` would swallow all
/// of a piped stdin and every call after it would answer `""` — the exact
/// silent-empty failure that made the 0.2 path put line splitting in the host.
///
/// `read-via-stream` is called again only when the buffer holds no newline, so
/// the piped case makes exactly one call and the interactive case makes one
/// per line.
pub fn emit_input(chunk: &mut Chunk, line: u32) {
    /// Stdin past the newline this call consumed. A GLOBAL, not a local: the
    /// surplus has to outlive the call that read it.
    const STDIN_BUF: &str = "__vybe_stdin_buf";

    let buf = chunk.alloc_scratch(1);
    let nl = chunk.alloc_scratch(1);
    let handle = chunk.alloc_scratch(1);
    let piece = chunk.alloc_scratch(1);
    let out = chunk.alloc_scratch(1);
    let len = chunk.alloc_scratch(1);

    let index_of = chunk.add_import("ecma:string", "indexOf");
    let substring = chunk.add_import("ecma:string", "substring");
    let str_len = chunk.add_import("ecma:string", "length");
    let char_code = chunk.add_import("ecma:string", "charCodeAt");
    let read_via = chunk.add_import("wasi:cli/stdin", "read-via-stream");
    let at = chunk.add_import("ecma:array", "at");
    let undef_test = chunk.add_import("wasm:js-undefined", "test");

    // buf = whatever the last call left over, or "" the first time round.
    crate::primitives::globals::emit_read(chunk, STDIN_BUF, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, buf, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op_u16(Op::LOCAL_GET, buf, line);
    chunk.emit_call(undef_test, 1, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_if(line);
    chunk.emit_string_const("", line);
    chunk.emit_op_u16(Op::LOCAL_SET, buf, line);
    chunk.emit_end(line);

    // Fill until a newline is in hand or stdin is spent.
    let filled = chunk.emit_block(line);
    let (fill, _) = chunk.emit_loop_s(line);

    chunk.emit_op_u16(Op::LOCAL_GET, buf, line);
    chunk.emit_string_const("\n", line);
    chunk.emit_call(index_of, 2, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, nl, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_GE_S, line);
    chunk.emit_br_if(1, line);

    // Element 0 of the tuple is the `stream<u8>` — an i32 readable handle,
    // per the canonical ABI.
    chunk.emit_call(read_via, 0, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_call(at, 2, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, handle, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_br_if(1, line);

    chunk.emit_op_u16(Op::LOCAL_GET, handle, line);
    emit_read_stream_to_string(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, piece, line);

    // A read that yielded nothing is EOF — going round again would spin
    // forever on a stdin that has no more to give.
    chunk.emit_op_u16(Op::LOCAL_GET, piece, line);
    chunk.emit_call(str_len, 1, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_LE_S, line);
    chunk.emit_br_if(1, line);

    chunk.emit_op_u16(Op::LOCAL_GET, buf, line);
    chunk.emit_op_u16(Op::LOCAL_GET, piece, line);
    super::strings::emit_concat(chunk, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, buf, line);

    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(fill);
    chunk.emit_end(line);
    chunk.patch_block(filled);

    // Split at the newline. Without one, stdin ended unterminated and what is
    // left IS the last line — answering "" there would drop it.
    chunk.emit_op_u16(Op::LOCAL_GET, buf, line);
    chunk.emit_string_const("\n", line);
    chunk.emit_call(index_of, 2, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, nl, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_LT_S, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, buf, line);
    chunk.emit_op_u16(Op::LOCAL_SET, out, line);
    chunk.emit_string_const("", line);
    chunk.emit_op_u16(Op::LOCAL_SET, buf, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, buf, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, nl, line);
    chunk.emit_call(substring, 3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, out, line);
    chunk.emit_op_u16(Op::LOCAL_GET, buf, line);
    chunk.emit_op_u16(Op::LOCAL_GET, nl, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_call(substring, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, buf, line);
    chunk.emit_end(line);

    // CRLF: the separator searched for is "\n", so on a Windows-terminated
    // stream the "\r" stays behind as the last character of the line.
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
    chunk.emit_call(str_len, 1, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, len, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_GT_S, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_call(char_code, 2, line);
    chunk.emit_i32_const(13, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_call(substring, 3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, out, line);
    chunk.emit_end(line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, buf, line);
    crate::primitives::globals::emit_write(chunk, STDIN_BUF, line);

    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
}

/// Emit readline using a pre-resolved import index.
pub fn emit_input_with_import(chunk: &mut Chunk, import_idx: u16, line: u32) {
    chunk.emit_call(import_idx, 0, line);
}

/// Emit print + input (prompt then read). Stack: [prompt_string] → [string]
pub fn emit_prompt_input(chunk: &mut Chunk, line: u32) {
    emit_print(chunk, 1, line);
    chunk.emit_op(Op::DROP, line);
    emit_input(chunk, line);
}

// ── File I/O lives in `fs_path.rs` ──────────────────────────────────────
//
// Seven emitters used to sit here — `emit_read_file`, `emit_open_file`,
// `emit_line_input` and friends — each a one-line `add_import("wasi:filesystem",
// "<verb>")` naming a function that is not in the WIT. All seven had ZERO
// callers; `platforms/wasi/src/fs.rs` said so in its own comment ("that helper
// has zero callers — it and its four `emit_*_file` neighbours are definitions
// only"). They are deleted rather than repointed: a dead shim that names a
// retired verb is how the next caller finds the wrong path.
//
// The live lowerings are `primitives::fs_path`, which composes only names the
// 0.3.1 WIT declares.

// ── Output buffering ────────────────────────────────────────────────────────
//
// A STACK of buffer frames, not a flag plus a "previous" slot. Nesting is the
// normal case (`ob_start()` inside a template inside a handler), and the depth
// is only known at runtime, so the representation has to be a real stack — the
// single-level shape could not express `ob_get_level() == 3` at all.
//
// One frame per active buffer, each carrying its own handler and options,
// because PHP's `ob_list_handlers()` / `ob_get_status(true)` report them
// per-level. Frames are Maps rather than a parallel set of arrays so that
// adding a field is a key, not a fifth global to keep in sync.

/// Global holding the buffer stack: an array of frames, innermost LAST.
/// Absent (null) until the first `ob_start` — see [`emit_ob_stack`].
const OB_STACK: &str = "__vybe_ob_stack";

/// Frame keys. Language-neutral: these are the fields of a buffer, not PHP's
/// spelling of them.
const OB_BUFFER: &str = "buffer";
const OB_HANDLER: &str = "handler";
const OB_CHUNK_SIZE: &str = "chunk_size";
const OB_FLAGS: &str = "flags";

fn global_get(chunk: &mut Chunk, key: &str, line: u32) {
    crate::primitives::globals::emit_read(chunk, key, line);
}

fn global_set(chunk: &mut Chunk, key: &str, line: u32) {
    crate::primitives::globals::emit_write(chunk, key, line);
}

/// Push the buffer stack, creating it on first use. Stack: [] → [array].
///
/// Lazily created rather than emitted as a module init so that a program which
/// never buffers pays nothing but the null check on its first write.
pub fn emit_ob_stack(chunks: &mut [Chunk], current: usize, line: u32) {
    global_get(&mut chunks[current], OB_STACK, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    super::collections::emit_array_new(chunks, current, 0, line);
    global_set(&mut chunks[current], OB_STACK, line);
    chunks[current].emit_end(line);
    global_get(&mut chunks[current], OB_STACK, line);
}

/// Number of active buffers. Stack: [] → [i32].
pub fn emit_ob_depth(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_ob_stack(chunks, current, line);
    super::collections::emit_array_length(&mut chunks[current], line);
}

/// The innermost frame. Stack: [] → [frame]. Callers must have established
/// depth > 0; there is no empty-stack frame to return.
pub fn emit_ob_top_frame(chunks: &mut [Chunk], current: usize, line: u32) {
    let stack_slot = chunks[current].alloc_scratch(1);
    emit_ob_stack(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, stack_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, stack_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stack_slot, line);
    super::collections::emit_array_length(&mut chunks[current], line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    super::collections::emit_get(chunks, current, line);
}

/// Read a field of the innermost frame. Stack: [] → [value].
pub fn emit_ob_top_field(chunks: &mut [Chunk], current: usize, field: &str, line: u32) {
    emit_ob_top_frame(chunks, current, line);
    chunks[current].emit_string_const(field, line);
    super::collections::emit_get(chunks, current, line);
}

/// Write a field of the innermost frame. Stack: [value] → [].
pub fn emit_ob_set_top_field(chunks: &mut [Chunk], current: usize, field: &str, line: u32) {
    let val_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, val_slot, line);

    emit_ob_top_frame(chunks, current, line);
    chunks[current].emit_string_const(field, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, val_slot, line);
    super::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

/// Write a string straight to stdout, bypassing any buffer. Stack: [] → [],
/// contents taken from `str_slot`.
pub fn emit_write_stdout_slot(chunk: &mut Chunk, str_slot: u16, line: u32) {
    emit_write_stream_slot(chunk, "wasi:cli/stdout", str_slot, line);
}

/// The same write against any `wasi:cli/*` sink that offers `write-via-stream`.
///
/// WASI 0.3 replaced the 0.2 pair `get-stdout()` + `[method]output-stream.
/// blocking-write-and-flush` with `write-via-stream(data: stream<u8>)`, because
/// a stream is now a Component Model TYPE rather than a WASI resource — the
/// whole `wasi:io` package is gone in 0.3. Anything still reaching for
/// `wasi:io/streams` is writing against a package that no longer exists.
pub fn emit_write_stream_slot(chunk: &mut Chunk, sink: &str, str_slot: u16, line: u32) {
    let write_idx = chunk.add_import(sink, "write-via-stream");
    let rd_slot = chunk.alloc_scratch(1);
    let wr_slot = chunk.alloc_scratch(1);
    emit_write_stdout_with_imports(chunk, write_idx, rd_slot, wr_slot, line, |chunk| {
        chunk.emit_op_u16(Op::LOCAL_GET, str_slot, line);
    });
}

/// Write a string to a FILE at an explicit position. Stack: `[] → []`.
///
/// `wasi:filesystem@0.3.1`:
///
///     write-via-stream: func(data: stream<u8>, offset: filesize)
///                       -> future<result<_, error-code>>
///
/// Same transport as stdout — the difference is only that the descriptor is
/// the receiver and the position is explicit. 0.2 had this inverted:
/// `write-via-stream(offset)` handed back an `output-stream` resource that the
/// guest then pushed into through `wasi:io/streams`. 0.3.1 takes the bytes and
/// answers a future, so the data flows the same direction as every other
/// write in this module.
///
/// `offset_slot` holds a `filesize`. Writing past the end is legal and extends
/// the file, zero-filling the gap (§write-via-stream) — which is what makes
/// record *n* at *n × width* expressible without a separate seek.
pub fn emit_write_descriptor_slot(
    chunk: &mut Chunk,
    desc_slot: u16,
    offset_slot: u16,
    str_slot: u16,
    line: u32,
) {
    let write_idx = chunk.add_import("wasi:filesystem/types", "[method]descriptor.write-via-stream");
    let rd_slot = chunk.alloc_scratch(1);
    let wr_slot = chunk.alloc_scratch(1);
    emit_write_via_stream(
        chunk,
        rd_slot,
        wr_slot,
        line,
        |chunk| {
            chunk.emit_op_u16(Op::LOCAL_GET, str_slot, line);
        },
        |chunk, rd| {
            chunk.emit_op_u16(Op::LOCAL_GET, desc_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, rd, line);
            chunk.emit_op_u16(Op::LOCAL_GET, offset_slot, line);
            chunk.emit_call(write_idx, 3, line);
            // The future is discarded here for the same reason stdout's is:
            // this is the fire-and-forget shape. A caller that needs the
            // `result<_, error-code>` keeps the future instead.
            chunk.emit_op(Op::DROP, line);
        },
    );
}

/// THE write. Stack: [string] → [].
///
/// Appends to the innermost buffer when one is active, otherwise goes to
/// stdout. Every language's print/echo/dump should route through this rather
/// than calling stdout directly — that is what makes output capture work
/// uniformly instead of per-builtin.
pub fn emit_write_or_buffer(chunks: &mut [Chunk], current: usize, line: u32) {
    let str_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, str_slot, line);

    emit_ob_depth(chunks, current, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);

    emit_ob_top_field(chunks, current, OB_BUFFER, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_slot, line);
    super::strings::emit_concat(&mut chunks[current], 2, line);
    emit_ob_set_top_field(chunks, current, OB_BUFFER, line);

    chunks[current].emit_else(line);
    emit_write_stdout_slot(&mut chunks[current], str_slot, line);
    chunks[current].emit_end(line);
}

/// Open a new buffer. Stack: [handler, chunk_size, flags] (per `argc`, missing
/// trailing args defaulted) → [true].
///
/// Always succeeds, matching PHP: `ob_start()` returns `false` only when a
/// handler refuses to start, which cannot happen for the default handler.
pub fn emit_ob_start(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    // Spill the supplied arguments; they arrive in source order so they come
    // off reversed.
    let flags_slot = chunks[current].alloc_scratch(1);
    let chunk_slot = chunks[current].alloc_scratch(1);
    let handler_slot = chunks[current].alloc_scratch(1);
    let slots = [handler_slot, chunk_slot, flags_slot];
    for i in (0..argc as usize).rev() {
        if i < slots.len() {
            chunks[current].emit_op_u16(Op::LOCAL_SET, slots[i], line);
        } else {
            chunks[current].emit_op(Op::DROP, line);
        }
    }
    for (i, slot) in slots.iter().enumerate() {
        if i >= argc as usize {
            chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, *slot, line);
        }
    }

    let frame_slot = chunks[current].alloc_scratch(1);
    super::collections::emit_map_new(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, frame_slot, line);

    for (field, slot) in [
        (OB_HANDLER, Some(handler_slot)),
        (OB_CHUNK_SIZE, Some(chunk_slot)),
        (OB_FLAGS, Some(flags_slot)),
        (OB_BUFFER, None),
    ] {
        chunks[current].emit_op_u16(Op::LOCAL_GET, frame_slot, line);
        chunks[current].emit_string_const(field, line);
        match slot {
            Some(s) => chunks[current].emit_op_u16(Op::LOCAL_GET, s, line),
            None => chunks[current].emit_string_const("", line),
        }
        super::collections::emit_set(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    }

    emit_ob_stack(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, frame_slot, line);
    super::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_bool_const(true, line);
}

/// Number of active buffers, as the numeric value languages report.
/// Stack: [] → [f64].
pub fn emit_ob_get_level(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_ob_depth(chunks, current, line);
    // Unsigned: a buffer count is never negative, so this is exact.
    chunks[current].emit_op(Op::F64_CONVERT_I32_U, line);
}

/// Emit `if depth > 0 { on_active } else { push false }`, leaving exactly one
/// value on the stack either way.
///
/// Every `ob_*` accessor has this shape — PHP returns `false` from all of them
/// when no buffer is active — so the branch is written once here rather than
/// re-derived (and mis-derived) per operation.
fn emit_when_buffering(
    chunks: &mut [Chunk],
    current: usize,
    line: u32,
    on_active: impl FnOnce(&mut [Chunk], usize),
) {
    let result_slot = chunks[current].alloc_scratch(1);
    emit_ob_depth(chunks, current, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    on_active(chunks, current);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// Contents of the innermost buffer, or `false`. Stack: [] → [string|false].
pub fn emit_ob_get_contents(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_when_buffering(chunks, current, line, |chunks, current| {
        emit_ob_top_field(chunks, current, OB_BUFFER, line);
    });
}

/// How a language measures the SIZE of buffered output. Stack: [string] → [num].
///
/// Output is bytes, but "length of a string" is not the same question in every
/// language — PHP's `strlen` counts UTF-8 bytes while the shared `emit_len`
/// counts code units, so `"éclair"` is 7 or 6 depending who asks. The language
/// supplies its own, exactly as it supplies `default_name`.
pub type LengthEmit = fn(&mut [Chunk], usize, u32);

fn emit_buffer_len(chunks: &mut [Chunk], current: usize, len: Option<LengthEmit>, line: u32) {
    match len {
        Some(emit) => emit(chunks, current, line),
        None => super::collections::emit_len(chunks, current, line),
    }
}

/// Size of the innermost buffer, or `false`. Stack: [] → [num|false].
pub fn emit_ob_get_length(
    chunks: &mut [Chunk],
    current: usize,
    len: Option<LengthEmit>,
    line: u32,
) {
    emit_when_buffering(chunks, current, line, |chunks, current| {
        emit_ob_top_field(chunks, current, OB_BUFFER, line);
        emit_buffer_len(chunks, current, len, line);
    });
}

/// Discard the innermost buffer's contents, keeping it open. Stack: [] → [bool].
pub fn emit_ob_clean(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_when_buffering(chunks, current, line, |chunks, current| {
        chunks[current].emit_string_const("", line);
        emit_ob_set_top_field(chunks, current, OB_BUFFER, line);
        chunks[current].emit_bool_const(true, line);
    });
}

/// Pop the innermost frame and leave its buffer contents on the stack.
/// Stack: [] → [string]. Caller must have established depth > 0.
fn emit_ob_pop(chunks: &mut [Chunk], current: usize, line: u32) {
    let contents_slot = chunks[current].alloc_scratch(1);
    emit_ob_top_field(chunks, current, OB_BUFFER, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, contents_slot, line);
    emit_ob_stack(chunks, current, line);
    super::collections::emit_pop(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, contents_slot, line);
}

/// Close the innermost buffer, discarding its contents. Stack: [] → [bool].
pub fn emit_ob_end_clean(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_when_buffering(chunks, current, line, |chunks, current| {
        emit_ob_pop(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_bool_const(true, line);
    });
}

/// Pop the innermost frame and write its contents outward, through the frame's
/// handler if it has one. Stack: [] → [raw_contents].
///
/// The RAW buffer is what comes back, while the HANDLED text is what gets
/// written — `ob_get_flush()` returns the former and outputs the latter, so the
/// two must not be conflated. The write happens after the pop, so
/// `emit_write_or_buffer` naturally targets the enclosing buffer (or stdout at
/// depth 0), which is where a flushed buffer's contents belong.
fn emit_ob_pop_and_flush(chunks: &mut [Chunk], current: usize, line: u32) {
    let handler_slot = chunks[current].alloc_scratch(1);
    let raw_slot = chunks[current].alloc_scratch(1);

    emit_ob_top_field(chunks, current, OB_HANDLER, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, handler_slot, line);
    emit_ob_pop(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, raw_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, handler_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, raw_slot, line);
    emit_write_or_buffer(chunks, current, line);
    chunks[current].emit_else(line);
    // [func_ref, receiver?, buffer] then CALL_REF — one user argument, the
    // receiver ahead of it being §10.2.1's argument 0.
    let recv =
        crate::primitives::callable::push_callback_from_slot(chunks, current, handler_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, raw_slot, line);
    crate::primitives::callable::emit_direct_invoke_chunk(&mut chunks[current], 1 + recv, line);
    emit_write_or_buffer(chunks, current, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, raw_slot, line);
}

/// Close the innermost buffer and write its contents to the next target out —
/// the enclosing buffer if there is one, otherwise stdout. Stack: [] → [bool].
pub fn emit_ob_end_flush(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_when_buffering(chunks, current, line, |chunks, current| {
        emit_ob_pop_and_flush(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_bool_const(true, line);
    });
}

/// Write the innermost buffer outward and empty it, WITHOUT closing it.
/// Stack: [] → [bool].
///
/// The frame is popped, written past, then pushed back — that is what makes the
/// write land in the ENCLOSING buffer rather than in the frame being flushed,
/// which is where a flush is supposed to go. The Map is a reference, so the
/// frame that comes back is the same one.
pub fn emit_ob_flush(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_when_buffering(chunks, current, line, |chunks, current| {
        let frame_slot = chunks[current].alloc_scratch(1);
        let handler_slot = chunks[current].alloc_scratch(1);
        let contents_slot = chunks[current].alloc_scratch(1);

        emit_ob_top_frame(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, frame_slot, line);
        emit_ob_top_field(chunks, current, OB_HANDLER, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, handler_slot, line);
        emit_ob_top_field(chunks, current, OB_BUFFER, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, contents_slot, line);

        chunks[current].emit_string_const("", line);
        emit_ob_set_top_field(chunks, current, OB_BUFFER, line);

        emit_ob_stack(chunks, current, line);
        super::collections::emit_pop(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, handler_slot, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_if(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, contents_slot, line);
        emit_write_or_buffer(chunks, current, line);
        chunks[current].emit_else(line);
        let recv = crate::primitives::callable::push_callback_from_slot(
            chunks, current, handler_slot, line,
        );
        chunks[current].emit_op_u16(Op::LOCAL_GET, contents_slot, line);
        crate::primitives::callable::emit_direct_invoke_chunk(&mut chunks[current], 1 + recv, line);
        emit_write_or_buffer(chunks, current, line);
        chunks[current].emit_end(line);

        emit_ob_stack(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, frame_slot, line);
        super::collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);

        chunks[current].emit_bool_const(true, line);
    });
}

/// Flush every still-open buffer, innermost first. Stack: [] → [].
///
/// Emitted at the end of a program: an unclosed buffer is flushed, not thrown
/// away. Each iteration pops one frame and writes it outward, so the innermost
/// contents land in its enclosing buffer and the outermost reach stdout.
pub fn emit_ob_flush_all(chunks: &mut [Chunk], current: usize, line: u32) {
    let block_p = chunks[current].emit_block(line);
    let (loop_p, _) = chunks[current].emit_loop_s(line);
    emit_ob_depth(chunks, current, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_LE_S, line);
    chunks[current].emit_br_if(1, line);
    emit_ob_pop_and_flush(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_p);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block_p);
}

/// Build the status record for the frame at `index_slot`. Stack: [] → [map].
fn emit_ob_status_record(
    chunks: &mut [Chunk],
    current: usize,
    stack_slot: u16,
    index_slot: u16,
    default_name: &str,
    len: Option<LengthEmit>,
    line: u32,
) {
    let frame_slot = chunks[current].alloc_scratch(1);
    let out_slot = chunks[current].alloc_scratch(1);
    let name_slot = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_GET, stack_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    super::collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, frame_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, frame_slot, line);
    chunks[current].emit_string_const(OB_HANDLER, line);
    super::collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, name_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const(default_name, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, name_slot, line);
    chunks[current].emit_end(line);

    super::collections::emit_map_new(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);

    // `level` is the frame's index; `buffer_used` its current byte count. The
    // rest are the fixed fields the status record is documented to carry.
    for field in [
        "name",
        "type",
        "flags",
        "level",
        "chunk_size",
        "buffer_size",
        "buffer_used",
    ] {
        chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
        chunks[current].emit_string_const(field, line);
        match field {
            "name" => chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line),
            "level" => {
                chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
                chunks[current].emit_op(Op::F64_CONVERT_I32_U, line);
            }
            "chunk_size" => {
                chunks[current].emit_op_u16(Op::LOCAL_GET, frame_slot, line);
                chunks[current].emit_string_const(OB_CHUNK_SIZE, line);
                super::collections::emit_get(chunks, current, line);
            }
            "flags" => {
                chunks[current].emit_op_u16(Op::LOCAL_GET, frame_slot, line);
                chunks[current].emit_string_const(OB_FLAGS, line);
                super::collections::emit_get(chunks, current, line);
            }
            "buffer_used" => {
                chunks[current].emit_op_u16(Op::LOCAL_GET, frame_slot, line);
                chunks[current].emit_string_const(OB_BUFFER, line);
                super::collections::emit_get(chunks, current, line);
                emit_buffer_len(chunks, current, len, line);
            }
            _ => chunks[current].emit_f64_const(0.0, line),
        }
        super::collections::emit_set(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

/// Buffer status. Stack: [] → [map] for the innermost buffer (empty when none),
/// or [array-of-maps], outermost first, when `full` is requested.
pub fn emit_ob_get_status(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    default_name: &str,
    len: Option<LengthEmit>,
    line: u32,
) {
    let full_slot = chunks[current].alloc_scratch(1);
    if argc >= 1 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, full_slot, line);
        for _ in 1..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
    } else {
        chunks[current].emit_bool_const(false, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, full_slot, line);
    }

    let stack_slot = chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].alloc_scratch(1);
    let i_slot = chunks[current].alloc_scratch(1);
    emit_ob_stack(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, stack_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, full_slot, line);
    super::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);

    // Full: one record per frame, outermost first.
    super::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    let block_p = chunks[current].emit_block(line);
    let (loop_p, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stack_slot, line);
    super::collections::emit_array_length(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    emit_ob_status_record(chunks, current, stack_slot, i_slot, default_name, len, line);
    super::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_p);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block_p);

    chunks[current].emit_else(line);

    // Innermost only — an empty map when nothing is buffering.
    chunks[current].emit_op_u16(Op::LOCAL_GET, stack_slot, line);
    super::collections::emit_array_length(&mut chunks[current], line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stack_slot, line);
    super::collections::emit_array_length(&mut chunks[current], line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    emit_ob_status_record(chunks, current, stack_slot, i_slot, default_name, len, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);
    super::collections::emit_map_new(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// Names of the active handlers, innermost LAST. Stack: [] → [array].
///
/// `default_name` is what a frame with no handler reports. The shared primitive
/// has no word for "no handler" — PHP spells it `"default output handler"` —
/// so the language supplies its own.
pub fn emit_ob_list_handlers(chunks: &mut [Chunk], current: usize, default_name: &str, line: u32) {
    let stack_slot = chunks[current].alloc_scratch(1);
    let out_slot = chunks[current].alloc_scratch(1);
    let i_slot = chunks[current].alloc_scratch(1);

    emit_ob_stack(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, stack_slot, line);
    super::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);

    let block_p = chunks[current].emit_block(line);
    let (loop_p, _) = chunks[current].emit_loop_s(line);
    // Break once the index reaches the depth.
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stack_slot, line);
    super::collections::emit_array_length(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    let handler_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stack_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    super::collections::emit_get(chunks, current, line);
    chunks[current].emit_string_const(OB_HANDLER, line);
    super::collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, handler_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, handler_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const(default_name, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, handler_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, handler_slot, line);
    super::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_p);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block_p);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

/// Contents of the innermost buffer, then close it discarding. Stack: []
/// → [string|false].
pub fn emit_ob_get_clean(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_when_buffering(chunks, current, line, |chunks, current| {
        // The handler runs on a clean as well as on a flush — it is how a
        // throwing handler surfaces rather than being swallowed — and its
        // result is what comes back. Nothing is WRITTEN: that is the whole
        // difference between `ob_get_clean` and `ob_get_flush`.
        let handler_slot = chunks[current].alloc_scratch(1);
        emit_ob_top_field(chunks, current, OB_HANDLER, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, handler_slot, line);

        let raw_slot = chunks[current].alloc_scratch(1);
        emit_ob_pop(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, raw_slot, line);

        let out_slot = chunks[current].alloc_scratch(1);
        chunks[current].emit_op_u16(Op::LOCAL_GET, handler_slot, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_if(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, raw_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
        chunks[current].emit_else(line);
        let recv = crate::primitives::callable::push_callback_from_slot(
            chunks, current, handler_slot, line,
        );
        chunks[current].emit_op_u16(Op::LOCAL_GET, raw_slot, line);
        crate::primitives::callable::emit_direct_invoke_chunk(&mut chunks[current], 1 + recv, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
        chunks[current].emit_end(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    });
}

/// Contents of the innermost buffer, then close it flushing outward.
/// Stack: [] → [string|false].
pub fn emit_ob_get_flush(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_when_buffering(chunks, current, line, |chunks, current| {
        emit_ob_pop_and_flush(chunks, current, line);
    });
}


// ── Linkable chunk builders ──────────────────────────────────────────────────
//
// Linkable chunk builders — the standalone-chunk packaging of what the
// `emit_*` forms splice inline. A language prefix in a name records which
// frontend first needed a linkable chunk, not a language-specific meaning.

// `build_pascal_write` / `build_pascal_writeln` removed — nothing
// referenced `__vybe_pascal_write{,ln}`. Pascal `Write`/`WriteLn` route
// through the buffered writer in this module instead. Their shared
// `emit_pascal_write_buffer` (read the buffer global, fold null and
// undefined to `""`) went with them as its only two callers.
