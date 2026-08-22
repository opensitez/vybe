//! WASM Exception Handling proposal compliance tests (final/exnref phase).
//!
//! Spec source: `proposals/*/proposals/exception-handling/Exceptions.md`.
//! These tests encode the PROPOSAL's semantics, not the VM's current
//! behavior — where the VM deviates they MUST fail until it is fixed:
//!
//!   1. Tags are module entities "created fresh each time" — two tag
//!      declarations are DISTINCT even with identical names/signatures.
//!      (Current: `add_exception_tag` dedups by name.)
//!   2. `throw` takes a TAG INDEX immediate and packages the payload with
//!      that tag. (Current: `THROW` has no immediate.)
//!   3. Catch clauses match by TAG IDENTITY ONLY — never by inspecting
//!      the payload. (Current: the VM string-matches the clause's tag
//!      NAME against the payload's `__type`/`__types` stamps, including
//!      subtype matching, which does not exist in WASM.)
//!   4. Clause kinds: `catch tag l` / `catch_ref tag l` / `catch_all l` /
//!      `catch_all_ref l`. `catch` pushes the payload; `catch_all` pushes
//!      NOTHING; the `_ref` forms additionally push an `exnref` that
//!      `throw_ref` rethrows. (Current: single catch-all form that always
//!      pushes the payload value; no exnref.)
//!   5. "If no catch clauses were matched, the exception is implicitly
//!      rethrown" (propagates to the enclosing try).
//!   6. Traps are NOT caught by `try_table`.
//!
//! # Internal fixed-width encoding under test
//!
//! Internal opcodes are fixed-width (4-byte op, u16 operands); WASM's
//! LEB variable-width lives only at the serialization boundary. The
//! spec-shaped internal encoding these tests emit:
//!
//! ```text
//! TRY_TABLE: op | u8 clause_count | per clause:
//!            u8 kind (0=catch 1=catch_ref 2=catch_all 3=catch_all_ref)
//!            u16 tag_idx (0 and ignored for catch_all/_ref)
//!            u16 offset  (forward, relative to the end of this clause)
//! THROW:     op | u16 tag_idx   (pops payload per the tag's arity)
//! THROW_REF: op                 (pops exnref, rethrows it)
//! END:       spec block-end (0x0B); closing a try_table's `end` removes its
//!            handler group via the `is_try` label (replaces retired TRY_END)
//! ```

use std::sync::atomic::{AtomicUsize, Ordering};
use std::sync::{Arc, Mutex};
use vybe_runtime::value::Object;
use vybe_runtime::{Chunk, Op, VM, Value};

/// Unique names for test-argument globals, so reused VMs never collide.
static TEST_GLOBAL_SEQ: AtomicUsize = AtomicUsize::new(0);

const KIND_CATCH: u8 = 0;
const KIND_CATCH_REF: u8 = 1;
const KIND_CATCH_ALL: u8 = 2;
const KIND_CATCH_ALL_REF: u8 = 3;

/// Marks where one clause's handler begins. Returned by [`emit_try_table`],
/// consumed by [`patch_clause`]; opaque on purpose, so a test never spells the
/// encoding out.
#[derive(Clone, Copy)]
struct ClauseTok {
    /// The first clause's handler also closes the `try_table` body.
    closes_try: bool,
}

/// How many values a clause of this kind delivers, for tags of arity 1 (which
/// is every tag these tests declare). The handler block must carry this as its
/// result arity: the branch keeps exactly that many values.
fn clause_arity(kind: u8) -> u8 {
    match kind {
        KIND_CATCH => 1,          // payload
        KIND_CATCH_REF => 2,      // payload + exnref
        KIND_CATCH_ALL => 0,      // spec: no values
        KIND_CATCH_ALL_REF => 1,  // exnref only
        _ => 0,
    }
}

/// Open a spec try region: one HANDLER BLOCK per clause, then the `try_table`.
///
/// Routed through `Chunk::emit_try_table_clauses` — the documented single
/// source of truth — rather than re-spelling the byte layout here. A test that
/// hand-encodes the layout asserts what the encoder does, so it cannot detect a
/// wrong encoder; that is precisely how the `labelidx`-as-byte-offset defect
/// survived in this file.
///
/// Blocks are emitted in REVERSE so clause 0's block is innermost and is
/// therefore `labelidx 0`. Block `i`'s `end` is where handler `i` begins.
fn emit_try_table(c: &mut Chunk, clauses: &[(u8, u16)]) -> Vec<ClauseTok> {
    for (kind, _) in clauses.iter().rev() {
        c.emit_block_typed(0, clause_arity(*kind));
    }
    let triples: Vec<(u8, u16, u16)> = clauses
        .iter()
        .enumerate()
        .map(|(i, &(kind, tag))| (kind, tag, i as u16))
        .collect();
    // Blocktype: these tests' try bodies produce no values.
    c.emit_try_table_clauses(0, 0, &triples, 0);
    (0..clauses.len())
        .map(|i| ClauseTok { closes_try: i == 0 })
        .collect()
}

/// Begin a clause's handler here.
///
/// Nothing is patched: the clause names a `labelidx`, so the handler's position
/// is decided by BLOCK STRUCTURE. Closing the clause's block is what places the
/// handler — its `end` is the branch target. Call these in clause order; clause
/// 0's block is innermost, so it closes first.
fn patch_clause(c: &mut Chunk, tok: ClauseTok) {
    if tok.closes_try {
        c.emit_op(Op::END, 0); // close the try_table body
    }
    c.emit_op(Op::END, 0); // close this clause's handler block
}

/// Spec `throw <tagidx>` — payload must already be on the stack.
fn emit_throw(c: &mut Chunk, tag: u16) {
    c.emit_op(Op::THROW, 0);
    c.emit((tag >> 8) as u8, 0);
    c.emit((tag & 0xff) as u8, 0);
}

fn push_str(c: &mut Chunk, s: &str) {
    c.emit_string_const(s, 0);
}

fn ret(c: &mut Chunk) {
    c.emit_op(Op::RETURN, 0);
}

fn run(c: Chunk) -> Result<Value, String> {
    VM::new().run(vec![c]).map_err(|e| e.to_string())
}

fn s(v: &Value) -> String {
    format!("{v}")
}

// ─────────────────────────────────────────────────────────────────────
// §"Exception tags": declarations create FRESH entities
// ─────────────────────────────────────────────────────────────────────

/// "The tag section is a list of declared tags that are created fresh."
/// Declaring two tags yields two distinct tag indices — even when the
/// debug names collide. Tag identity is the declaration, not the name.
#[test]
fn tag_declarations_are_fresh_entities() {
    let mut c = Chunk::new("<script>");
    let t1 = c.declare_exception_tag("E", 1);
    let t2 = c.declare_exception_tag("E", 1);
    assert_ne!(
        t1, t2,
        "each tag declaration must create a fresh entity (spec tag section); \
         name-keyed deduplication is not tag identity"
    );
}

// ─────────────────────────────────────────────────────────────────────
// §"Throwing an exception" + §"Try blocks": identity match + payload
// ─────────────────────────────────────────────────────────────────────

/// `throw $t` inside `try_table (catch $t L)` transfers to L with the
/// tag's payload values on the stack.
#[test]
fn catch_matches_thrown_tag_and_delivers_payload() {
    let mut c = Chunk::new("<script>");
    let t = c.declare_exception_tag("E", 1);
    let patches = emit_try_table(&mut c, &[(KIND_CATCH, t)]);
    push_str(&mut c, "boom");
    emit_throw(&mut c, t);
    // unreachable fall-through
    push_str(&mut c, "not-thrown");
    ret(&mut c);
    patch_clause(&mut c, patches[0]);
    // handler: [payload]
    ret(&mut c);

    let v = run(c).expect("caught by matching tag");
    assert_eq!(s(&v), "boom");
}

/// Catch clauses match by TAG IDENTITY, never by payload inspection.
/// The payload here is stamped with `__type`/`__types` values that the
/// legacy matcher would name/subtype-match against the clause's tag
/// name ("Exception") — per spec this must NOT match: the thrown tag is
/// a different entity. The enclosing `catch $tB` receives it instead.
#[test]
fn catch_ignores_payload_stamps_tag_identity_only() {
    let mut c = Chunk::new("<script>");
    let t_exception = c.declare_exception_tag("Exception", 1);
    let t_type_error = c.declare_exception_tag("TypeError", 1);

    // Payload deliberately stamped to bait name/subtype matching.
    let mut obj = Object::new();
    obj.properties
        .insert("__type".into(), Value::String(Arc::from("TypeError")));
    obj.properties.insert(
        "__types".into(),
        Value::Object(Arc::new(Mutex::new(Object::new_array(vec![
            Value::String(Arc::from("TypeError")),
            Value::String(Arc::from("Exception")),
        ])))),
    );
    obj.properties
        .insert("message".into(), Value::String(Arc::from("stamped")));
    let payload = Value::Object(Arc::new(Mutex::new(obj)));
    let payload_global = format!(
        "__test_arg_{}",
        TEST_GLOBAL_SEQ.fetch_add(1, Ordering::Relaxed)
    );

    let outer = emit_try_table(&mut c, &[(KIND_CATCH, t_type_error)]);
    let inner = emit_try_table(&mut c, &[(KIND_CATCH, t_exception)]);
    let payload_ci = c.intern_string_constant(&payload_global);
    c.emit_op_u16(Op::GLOBAL_GET, payload_ci, 0);
    emit_throw(&mut c, t_type_error);
    push_str(&mut c, "not-thrown");
    ret(&mut c);
    patch_clause(&mut c, inner[0]);
    // Legacy matcher lands here: clause tag NAME "Exception" appears in
    // the payload's __types chain. Spec: tag identity differs → no match.
    push_str(&mut c, "WRONG: matched by payload stamps");
    ret(&mut c);
    patch_clause(&mut c, outer[0]);
    // Spec path: implicit rethrow reaches the enclosing catch $tTypeError.
    push_str(&mut c, "outer-caught-by-identity");
    ret(&mut c);

    let mut vm = VM::new();
    vm.set_global_owned(payload_global, payload);
    let v = vm
        .run(vec![c])
        .map_err(|e| e.to_string())
        .expect("outer typed clause catches by identity");
    assert_eq!(s(&v), "outer-caught-by-identity");
}

/// Two tags with the SAME signature are still distinct entities: a
/// clause for one must not catch a throw of the other.
#[test]
fn distinct_tags_with_same_signature_do_not_match() {
    let mut c = Chunk::new("<script>");
    let t1 = c.declare_exception_tag("E", 1);
    let t2 = c.declare_exception_tag("E", 1);

    let patches = emit_try_table(&mut c, &[(KIND_CATCH, t1)]);
    push_str(&mut c, "boom");
    emit_throw(&mut c, t2);
    push_str(&mut c, "not-thrown");
    ret(&mut c);
    patch_clause(&mut c, patches[0]);
    push_str(&mut c, "WRONG: t1 clause caught t2 throw");
    ret(&mut c);

    let err = run(c).expect_err("uncaught: no clause matches tag t2");
    assert!(
        !err.contains("WRONG"),
        "clause for t1 must not catch a throw of distinct tag t2"
    );
}

// ─────────────────────────────────────────────────────────────────────
// §"Try blocks": clause kinds and ordering
// ─────────────────────────────────────────────────────────────────────

/// `catch_all` catches any tag but "in case of catch_all and
/// catch_all_ref, no values are pushed" — the handler must NOT receive
/// the payload. A sentinel pushed before the try must be TOS in the
/// handler.
#[test]
fn catch_all_matches_any_tag_but_delivers_no_payload() {
    let mut c = Chunk::new("<script>");
    let t = c.declare_exception_tag("E", 1);
    push_str(&mut c, "sentinel"); // must still be TOS inside the handler
    let patches = emit_try_table(&mut c, &[(KIND_CATCH_ALL, 0)]);
    push_str(&mut c, "payload");
    emit_throw(&mut c, t);
    push_str(&mut c, "not-thrown");
    ret(&mut c);
    patch_clause(&mut c, patches[0]);
    // handler: catch_all pushes NOTHING → TOS is the sentinel
    ret(&mut c);

    let v = run(c).expect("catch_all catches any tag");
    assert_eq!(
        s(&v),
        "sentinel",
        "spec catch_all pushes no values; the payload must not appear"
    );
}

/// "Catch clauses are tried in the order they appear ... until one
/// matches." A matching typed clause listed first wins over a later
/// catch_all; a catch_all listed first wins over a later typed clause.
#[test]
fn clauses_are_tried_in_order_first_match_wins() {
    // typed first
    let mut c = Chunk::new("<script>");
    let t = c.declare_exception_tag("E", 1);
    let patches = emit_try_table(&mut c, &[(KIND_CATCH, t), (KIND_CATCH_ALL, 0)]);
    push_str(&mut c, "boom");
    emit_throw(&mut c, t);
    ret(&mut c);
    patch_clause(&mut c, patches[0]);
    ret(&mut c); // typed handler returns the payload "boom"
    patch_clause(&mut c, patches[1]);
    push_str(&mut c, "WRONG: catch_all ran before matching catch");
    ret(&mut c);
    let v = run(c).expect("typed clause first");
    assert_eq!(s(&v), "boom");

    // catch_all first
    let mut c = Chunk::new("<script>");
    let t = c.declare_exception_tag("E", 1);
    push_str(&mut c, "all-first");
    let patches = emit_try_table(&mut c, &[(KIND_CATCH_ALL, 0), (KIND_CATCH, t)]);
    push_str(&mut c, "boom");
    emit_throw(&mut c, t);
    ret(&mut c);
    patch_clause(&mut c, patches[0]);
    ret(&mut c); // catch_all pushes nothing → returns "all-first"
    patch_clause(&mut c, patches[1]);
    ret(&mut c);
    let v = run(c).expect("catch_all first");
    assert_eq!(s(&v), "all-first");
}

// ─────────────────────────────────────────────────────────────────────
// §"Exception references": catch_ref / catch_all_ref / throw_ref
// ─────────────────────────────────────────────────────────────────────

/// `catch_ref $t L`: L receives the payload values AND an exnref on
/// top; `throw_ref` rethrows that exact exception (same tag, same
/// payload) to the enclosing try.
#[test]
fn catch_ref_delivers_exnref_and_throw_ref_rethrows() {
    let mut c = Chunk::new("<script>");
    let t = c.declare_exception_tag("E", 1);
    let exn_slot = 0u16;
    c.local_count = 1;

    let outer = emit_try_table(&mut c, &[(KIND_CATCH, t)]);
    let inner = emit_try_table(&mut c, &[(KIND_CATCH_REF, t)]);
    push_str(&mut c, "original-payload");
    emit_throw(&mut c, t);
    push_str(&mut c, "not-thrown");
    ret(&mut c);
    patch_clause(&mut c, inner[0]);
    // handler: [payload, exnref] — park exnref, drop payload, rethrow
    c.emit_op_u16(Op::LOCAL_SET, exn_slot, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_op_u16(Op::LOCAL_GET, exn_slot, 0);
    c.emit_op(Op::THROW_REF, 0);
    patch_clause(&mut c, outer[0]);
    // outer typed handler receives the ORIGINAL payload
    ret(&mut c);

    let v = run(c).expect("throw_ref rethrows to enclosing catch");
    assert_eq!(s(&v), "original-payload");
}

/// `catch_all_ref L`: L receives ONLY the exnref (no payload); the
/// rethrown exception keeps its identity — an enclosing typed clause
/// for the original tag catches it with the original payload.
#[test]
fn catch_all_ref_delivers_only_exnref_and_preserves_identity() {
    let mut c = Chunk::new("<script>");
    let t = c.declare_exception_tag("E", 1);

    let outer = emit_try_table(&mut c, &[(KIND_CATCH, t)]);
    let inner = emit_try_table(&mut c, &[(KIND_CATCH_ALL_REF, 0)]);
    push_str(&mut c, "payload-x");
    emit_throw(&mut c, t);
    push_str(&mut c, "not-thrown");
    ret(&mut c);
    patch_clause(&mut c, inner[0]);
    // handler: [exnref] only — rethrow it directly
    c.emit_op(Op::THROW_REF, 0);
    patch_clause(&mut c, outer[0]);
    ret(&mut c);

    let v = run(c).expect("identity preserved through catch_all_ref/throw_ref");
    assert_eq!(s(&v), "payload-x");
}

// ─────────────────────────────────────────────────────────────────────
// §"Throwing an exception": propagation
// ─────────────────────────────────────────────────────────────────────

/// "If no catch clauses were matched, the exception is implicitly
/// rethrown" — an inner try whose only clause is for a different tag
/// is transparent; the enclosing matching clause catches.
#[test]
fn unmatched_exception_implicitly_rethrown_to_enclosing_try() {
    let mut c = Chunk::new("<script>");
    let t_a = c.declare_exception_tag("A", 1);
    let t_b = c.declare_exception_tag("B", 1);

    let outer = emit_try_table(&mut c, &[(KIND_CATCH, t_a)]);
    let inner = emit_try_table(&mut c, &[(KIND_CATCH, t_b)]);
    push_str(&mut c, "escapes-inner");
    emit_throw(&mut c, t_a);
    push_str(&mut c, "not-thrown");
    ret(&mut c);
    patch_clause(&mut c, inner[0]);
    push_str(&mut c, "WRONG: inner B clause caught an A throw");
    ret(&mut c);
    patch_clause(&mut c, outer[0]);
    ret(&mut c);

    let v = run(c).expect("outer catches after inner is transparent");
    assert_eq!(s(&v), "escapes-inner");
}

/// A throw with no matching clause anywhere escapes the instance as a
/// runtime error (the embedder surface of an uncaught exception).
#[test]
fn uncaught_exception_escapes_as_runtime_error() {
    let mut c = Chunk::new("<script>");
    let t = c.declare_exception_tag("E", 1);
    push_str(&mut c, "nobody-catches");
    emit_throw(&mut c, t);
    ret(&mut c);

    run(c).expect_err("uncaught exception must surface as an error");
}

/// The structural `end` that closes a `try_table` block deactivates the
/// handler: a throw AFTER the try block must not be caught by it. (This is
/// the spec block-end mechanism that replaced the retired custom TRY_END.)
#[test]
fn try_end_deactivates_handlers() {
    let mut c = Chunk::new("<script>");
    let t = c.declare_exception_tag("E", 1);
    let patches = emit_try_table(&mut c, &[(KIND_CATCH, t)]);
    // empty try body
    c.emit_op(Op::END, 0);
    push_str(&mut c, "late");
    emit_throw(&mut c, t); // outside the try — must NOT be caught
    ret(&mut c);
    patch_clause(&mut c, patches[0]);
    push_str(&mut c, "WRONG: handler caught a throw after try_end");
    ret(&mut c);

    let err = run(c).expect_err("throw after try_end is uncaught");
    assert!(!err.contains("WRONG"));
}

// ─────────────────────────────────────────────────────────────────────
// §"Traps": not catchable
// ─────────────────────────────────────────────────────────────────────

/// "The try_table instruction does not catch exceptions generated from
/// traps." An `unreachable` trap inside the try body must escape even a
/// catch_all.
#[test]
fn traps_are_not_caught_by_try_table() {
    let mut c = Chunk::new("<script>");
    let patches = emit_try_table(&mut c, &[(KIND_CATCH_ALL, 0)]);
    c.emit_op(Op::UNREACHABLE, 0);
    ret(&mut c);
    patch_clause(&mut c, patches[0]);
    push_str(&mut c, "WRONG: trap was caught");
    ret(&mut c);

    let err = run(c).expect_err("traps escape try_table");
    assert!(
        err.contains("unreachable"),
        "the trap itself must surface, got: {err}"
    );
}
