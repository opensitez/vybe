//! Exception handling helpers — shared try/catch/finally bytecode patterns.
//!
//! All compilers emit the same opcodes for exception handling:
//! - try_table (real WASM EH Phase 4) → body → try_end → handler
//! - try_end pops the handler on normal (non-throwing) exit

use std::sync::Arc;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

use crate::reflection;

/// Build a standard exception constructor chunk.
/// All languages should use this shape: { __type, __exception_type, name, message }.
/// This ensures Python `except ValueError` can catch a Dart `throw ValueError("...")`.
pub fn emit_exception_constructor(
    chunk: &mut Chunk,
    this_slot: u16,
    exc_name: &str,
    msg_slot: u16,
    line: u32,
) {
    // Create object
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, this_slot, line);

    // Shared type/reflection stamps. `__exception_type` remains for older
    // language surfaces, while reflection consumers read `__typename`/`__kind`.
    let canon = canonical_exception_name(exc_name);
    for (key, val) in [
        (reflection::FIELD_TYPE, canon),
        (reflection::FIELD_TYPE_NAME, canon),
        ("__exception_type", canon),
        (
            reflection::FIELD_KIND,
            reflection::ReflectKind::Exception.as_str(),
        ),
    ] {
        chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
        chunk.emit_string_const(val, line);
        let k = chunk.add_constant(Value::String(Arc::from(key)));
        chunk.emit_op_u16(Op::STRUCT_SET, k, line);
        chunk.emit_op(Op::DROP, line);
    }

    // name = exc_name (JS Error convention)
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_string_const(exc_name, line);
    let n_key = chunk.add_constant(Value::String(Arc::from("name")));
    chunk.emit_op_u16(Op::STRUCT_SET, n_key, line);
    chunk.emit_op(Op::DROP, line);

    // message = msg_slot
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, msg_slot, line);
    let m_key = chunk.add_constant(Value::String(Arc::from("message")));
    chunk.emit_op_u16(Op::STRUCT_SET, m_key, line);
    chunk.emit_op(Op::DROP, line);
}

/// §20.5: finish a JS error instance minted by an `ecma:error.*`
/// constructor — link [[Prototype]] to the canonical ctor's `prototype`
/// (wired by the JS prelude on the `__ctor_<Kind>` anchor) and drop the
/// host's own `name` stamp: per spec `name` is a prototype property
/// (`new Error("x").hasOwnProperty("name")` is false), and with the
/// chain in place it resolves through the prototype.
/// Stack: [err] → [err]
pub fn emit_finish_js_error_instance(chunk: &mut Chunk, kind: &str, line: u32) {
    // err.__proto__ = <kind ctor>.prototype
    crate::instructions::core_wasm::dup(chunk, line); // [err, err]
    let ctor_key = chunk.add_constant(Value::String(Arc::from(format!("__ctor_{kind}").as_str())));
    chunk.emit_op_u16(Op::GLOBAL_GET, ctor_key, line); // [err, err, ctor]
    let proto_key = chunk.add_constant(Value::String(Arc::from("prototype")));
    chunk.emit_op_u16(Op::STRUCT_GET, proto_key, line); // [err, err, proto]
    let link_key = chunk.add_constant(Value::String(Arc::from("__proto__")));
    chunk.emit_op_u16(Op::STRUCT_SET, link_key, line); // [err, err]
    chunk.emit_op(Op::DROP, line); // [err]
    // delete err.name — own stamp off, prototype `name` takes over
    crate::instructions::core_wasm::dup(chunk, line); // [err, err]
    chunk.emit_string_const("name", line); // [err, err, "name"]
    let del = chunk.add_import("ecma:object", "delete");
    chunk.emit_call(del, 2, line); // [err, bool]
    chunk.emit_op(Op::DROP, line); // [err]
}

/// Standard exception type names shared across all languages.
/// Maps language-specific names to a canonical set.
pub fn canonical_exception_name(name: &str) -> &str {
    // Defensive: walkers occasionally include trailing whitespace from the
    // type span (e.g. C# `catch (Exception e)` produces "Exception "). Trim
    // before matching AND in the fallthrough so the runtime-side
    // `STRUCT_GET __exception_type` compare doesn't miss on a trailing
    // space mismatch.
    let trimmed = name.trim();
    let short_name = trimmed.rsplit('.').next().unwrap_or(trimmed).trim();
    match short_name.to_lowercase().as_str() {
        // Python → canonical
        "valueerror" | "formaterror" | "formatexception" => "ValueError",
        "typeerror" => "TypeError",
        "keyerror" | "keynotfoundexception" => "KeyError",
        "indexerror" | "indexoutofrangeexception" | "rangerror" => "IndexError",
        "runtimeerror" | "runtimeexception" => "RuntimeError",
        "stopiteration" | "stateexception" => "StopIteration",
        "attributeerror" | "nosuchmethoderror" => "AttributeError",
        "zerodivisionerror" | "integerdivisionbyzeroexception" => "ZeroDivisionError",
        "filenotfounderror" | "filenotfoundexception" => "FileNotFoundError",
        "importerror" => "ImportError",
        "notimplementederror" | "unimplementederror" => "NotImplementedError",
        "overflowerror" | "overflowexception" | "stackoverflowerror" => "OverflowError",
        "operationcanceledexception" | "taskcanceledexception" => "OperationCanceledException",
        "aggregateexception" => "AggregateException",
        "ioerror" | "ioexception" => "IOError",
        "oserror" => "OSError",
        "exception" | "error" => "Exception",
        _ => trimmed,
    }
}

/// Spec EH catch-clause kinds (exception-handling proposal `try_table`).
pub const CATCH_KIND_CATCH: u8 = 0;
pub const CATCH_KIND_CATCH_REF: u8 = 1;
pub const CATCH_KIND_CATCH_ALL: u8 = 2;
pub const CATCH_KIND_CATCH_ALL_REF: u8 = 3;

/// The shared language-exception tag (single-tag design, like every wasm
/// toolchain): payload arity 1 — the exception object. Imported by name so
/// all chunks/modules resolve to the SAME tag entity.
pub const EXCEPTION_TAG_NAME: &str = "vybe:exception";

/// Import the language-exception tag on `chunk` and return its tag index.
pub fn exception_tag(chunk: &mut Chunk) -> u16 {
    chunk.import_exception_tag(EXCEPTION_TAG_NAME, 1)
}

/// One catch clause of a `try_table`. `kind` is one of the `CATCH_KIND_*`
/// constants; `tag` selects the tag entity for `catch`/`catch_ref` and is
/// ignored for the `catch_all` kinds.
#[derive(Clone, Copy)]
pub struct TryTableClause {
    pub kind: u8,
    pub tag: u16,
}

/// Emit a `try_table` header with N catch clauses — the single source of
/// truth for the VM's internal try_table byte layout, shared by every
/// producer (the OO `emit_try_start` wrapper, the wast `WasmTryTable`
/// lowering, and the `.wasm` reader that imports foreign modules).
///
/// Internal fixed-width layout: `[try_table, u8 clause_count, per clause:
/// u8 kind, u16 tag, u16 offset]`. Clauses are matched by TAG IDENTITY in the
/// order given. Returns each clause's offset-placeholder position; patch it
/// with [`patch_catch`] once that clause's handler code has been emitted. The
/// caller emits the protected body next, then [`emit_try_end`].
/// Stack: unchanged.
pub fn emit_try_table(chunk: &mut Chunk, clauses: &[TryTableClause], line: u32) -> Vec<usize> {
    let pairs: Vec<(u8, u16)> = clauses.iter().map(|c| (c.kind, c.tag)).collect();
    chunk.emit_try_table_clauses(&pairs, line)
}

/// Emit the start of a try block. Returns the offset_pos to patch later.
/// Spec `try_table` with one `catch $vybe:exception` clause — the handler
/// receives the exception object (the tag's payload). Thin wrapper over the
/// shared [`emit_try_table`] primitive.
/// Stack: unchanged
pub fn emit_try_start(chunk: &mut Chunk, line: u32) -> usize {
    let tag = exception_tag(chunk);
    emit_try_table(
        chunk,
        &[TryTableClause {
            kind: CATCH_KIND_CATCH,
            tag,
        }],
        line,
    )[0]
}

/// Emit the structural `end` that closes the `try_table` block opened by
/// [`emit_try_start`]. Spec `try_table … end` IS a block: reaching this
/// `end` on normal completion pops the block's `is_try` label, which the
/// VM uses to also remove the exception-handler group (see the `END`
/// dispatch). This replaces the retired custom `TRY_END` opcode — the
/// caller must have accounted for the `try_table` block in its
/// `label_depth` (`+= 1` after `emit_try_start`, `-= 1` after this).
pub fn emit_try_end(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::END, line);
}

/// Patch the catch handler offset after the handler code has been emitted.
///
/// The VM reads `offset` (2 bytes) and computes `catch_ip = ip + offset`,
/// where ip is the position right after those 2 bytes (`offset_pos + 2`).
/// The forward distance from that ip to the current end of code is the offset.
pub fn patch_catch(chunk: &mut Chunk, offset_pos: usize) {
    let jump = chunk.current_offset() as i32 - (offset_pos as i32 + 2);
    chunk.code[offset_pos] = (jump >> 8) as u8;
    chunk.code[offset_pos + 1] = (jump & 0xff) as u8;
}

/// Emit a throw — takes the exception value from TOS.
/// Spec `throw <tagidx>`: the tag immediate selects the language-exception
/// tag; the object on the stack is its 1-ary payload.
/// Stack before: [exception_value]  Stack after: diverges
pub fn emit_throw(chunk: &mut Chunk, line: u32) {
    let tag = exception_tag(chunk);
    chunk.emit_op(Op::THROW, line);
    chunk.emit((tag >> 8) as u8, line);
    chunk.emit((tag & 0xff) as u8, line);
}

/// Returns true if `name` (case-insensitive) is one of the known
/// exception type names that should produce the canonical 4-field
/// shape via `emit_exception_new`. The list is the union of every
/// language's built-in exception types — adding a new language entry
/// only requires extending `canonical_exception_name`.
pub fn is_exception_type(name: &str) -> bool {
    let lower = name.to_lowercase();
    matches!(
        lower.as_str(),
        // Generic
        "exception" | "error" | "throwable"
        // Python / canonical
        | "valueerror" | "typeerror" | "keyerror" | "indexerror"
        | "runtimeerror" | "stopiteration" | "attributeerror"
        | "zerodivisionerror" | "filenotfounderror" | "importerror"
        | "notimplementederror" | "overflowerror" | "ioerror" | "oserror"
        | "baseexception" | "keyboardinterrupt" | "systemexit" | "generatorexit"
        | "lookuperror" | "unicodeerror" | "unicodedecodeerror" | "unicodeencodeerror"
        | "floatingpointerror" | "recursionerror" | "eoferror" | "memoryerror"
        | "buffererror" | "systemerror" | "modulenotfounderror" | "unboundlocalerror"
        | "indentationerror" | "taberror" | "stopasynciteration"
        | "permissionerror" | "timeouterror" | "fileexistserror"
        | "isadirectoryerror" | "notadirectoryerror" | "blockingioerror"
        | "brokenpipeerror" | "connectionerror" | "connectionreseterror"
        | "connectionrefusederror" | "connectionabortederror" | "processlookuperror"
        | "environmenterror" | "interruptederror" | "childprocesserror"
        // .NET / VB / C#
        | "systemexception" | "applicationexception" | "argumentexception" | "argumentnullexception"
        | "invalidoperationexception" | "notimplementedexception"
        | "notsupportedexception" | "nullreferenceexception"
        | "indexoutofrangeexception" | "keynotfoundexception"
        | "formatexception" | "stackoverflowerror" | "stackoverflowexception"
        | "integerdivisionbyzeroexception" | "rangerror" | "stateexception"
        | "filenotfoundexception" | "ioexception" | "formaterror"
        | "nosuchmethoderror" | "unimplementederror" | "overflowexception"
        | "operationcanceledexception" | "taskcanceledexception"
        // PHP
        | "runtimeexception" | "logicexception" | "domainexception"
        | "lengthexception" | "outofboundsexception" | "outofrangeexception"
        | "rangeexception" | "underflowexception"
        | "unexpectedvalueexception" | "invalidargumentexception"
        | "badfunctioncallexception" | "badmethodcallexception"
        | "arithmeticerror" | "compileerror" | "parseerror" | "assertionerror"
        | "jsonexception"
        | "unhandledmatcherror" | "divisionbyzeroerror" | "argumentcounterror"
        | "errorexception"
        // JS
        | "rangeerror" | "syntaxerror" | "referenceerror" | "urierror"
        | "evalerror" | "aggregateerror" | "suppressederror"
        // Ruby
        | "standarderror" | "argumenterror" | "nameerror" | "nomethoderror"
    )
}

/// Stack-based exception constructor. Use this in two phases:
///
/// 1. Caller emits `Op::STRUCT_NEW` and `emit_dup` to push `[obj, obj]`,
///    then emits the message expression to push `[obj, obj, msg]`.
/// 2. Caller invokes `emit_exception_new_finalize(chunk, exc_name, line)`
///    which consumes the inner `[obj, msg]` pair into `obj.message=msg`,
///    then stamps `__type`, `__exception_type` onto the outer obj.
///
/// Per ECMA-262 §20.5, name and constructor are inherited from Error.prototype,
/// not own properties, so they are not set here. JavaScript callers should ensure
/// proper prototype chain setup if needed.
///
/// Stack before: `[obj, obj, msg]`   Stack after: `[obj]`
///
/// Splitting the helper this way avoids the closure-vs-`&mut self`
/// borrow problem in language compilers (the compiler needs `&mut self`
/// to emit the message expression, which can't co-exist with a `&mut
/// chunk` borrow held by a closure-taking helper).
///
/// This is the **single source of truth** for `new SomeError(msg)` across
/// every language compiler. The name is normalized via
/// `canonical_exception_name` so PHP `RuntimeException`, Python
/// `RuntimeError`, JS `Error`, etc. all produce identical bytecode and
/// can therefore catch each other across language boundaries.
pub fn emit_exception_new_finalize(chunk: &mut Chunk, exc_name: &str, line: u32) {
    let canon = canonical_exception_name(exc_name);
    let original = exc_name.trim();

    // Coerce message to string per ECMA-262 §20.5.1.1 step 3
    let str_idx = chunk.add_import("ecma:string", "String");
    chunk.emit_call(str_idx, 1, line);
    // [obj, obj, msg_string] → [obj, msg_string] via struct_set "message"
    let m_key = chunk.add_constant(Value::String(Arc::from("message")));
    chunk.emit_op_u16(Op::STRUCT_SET, m_key, line);
    // [obj, msg_val] → [obj]
    chunk.emit_op(Op::DROP, line);

    // Shared type/reflection stamps use the canonical name for cross-language
    // catch dispatch and introspection compatibility.
    for (key, val) in [
        (reflection::FIELD_TYPE, canon),
        (reflection::FIELD_TYPE_NAME, canon),
        ("__exception_type", canon),
        (
            reflection::FIELD_KIND,
            reflection::ReflectKind::Exception.as_str(),
        ),
    ] {
        chunk.emit_dup(line);
        chunk.emit_string_const(val, line);
        let k = chunk.add_constant(Value::String(Arc::from(key)));
        chunk.emit_op_u16(Op::STRUCT_SET, k, line);
        chunk.emit_op(Op::DROP, line);
    }

    // Set name as a dynamic (non-indexed) property with the original language-specific name.
    // It will be added to __nonenum at the type level, making it non-enumerable.
    chunk.emit_dup(line);
    chunk.emit_string_const(original, line);
    let n_key = chunk.add_constant(Value::String(Arc::from("name")));
    chunk.emit_op_u16(Op::STRUCT_SET, n_key, line);
    chunk.emit_op(Op::DROP, line);
}

/// Canonical exception ancestor chains — the Python-shaped tree used by
/// non-`throwable_is_root` profiles (Python/.NET/Ruby). The first entry
/// is the type itself; each subsequent entry is a base class a typed
/// catch may name. `throwable_is_root` languages (PHP/Java, whose
/// `Error`/`Exception` are SIBLING branches) stamp their own chains and
/// must NOT use this table.
pub fn exception_ancestors(name: &str) -> &'static [&'static str] {
    match canonical_exception_name(name) {
        // BaseException-only branch (deliberately NOT under Exception)
        "KeyboardInterrupt" => &["KeyboardInterrupt", "BaseException"],
        "SystemExit" => &["SystemExit", "BaseException"],
        "GeneratorExit" => &["GeneratorExit", "BaseException"],
        "BaseException" => &["BaseException"],
        "Exception" => &["Exception", "BaseException"],
        // LookupError branch
        "KeyError" => &["KeyError", "LookupError", "Exception", "BaseException"],
        "IndexError" => &["IndexError", "LookupError", "Exception", "BaseException"],
        "LookupError" => &["LookupError", "Exception", "BaseException"],
        // ArithmeticError branch
        "ZeroDivisionError" => &[
            "ZeroDivisionError",
            "ArithmeticError",
            "Exception",
            "BaseException",
        ],
        "OverflowError" => &[
            "OverflowError",
            "ArithmeticError",
            "Exception",
            "BaseException",
        ],
        "FloatingPointError" => &[
            "FloatingPointError",
            "ArithmeticError",
            "Exception",
            "BaseException",
        ],
        "ArithmeticError" => &["ArithmeticError", "Exception", "BaseException"],
        // RuntimeError branch
        "NotImplementedError" => &[
            "NotImplementedError",
            "RuntimeError",
            "Exception",
            "BaseException",
        ],
        "RecursionError" => &[
            "RecursionError",
            "RuntimeError",
            "Exception",
            "BaseException",
        ],
        "RuntimeError" => &["RuntimeError", "Exception", "BaseException"],
        "OperationCanceledException" => &[
            "OperationCanceledException",
            "SystemException",
            "Exception",
            "BaseException",
        ],
        "AggregateException" => &[
            "AggregateException",
            "SystemException",
            "Exception",
            "BaseException",
        ],
        // OSError branch (IOError/EnvironmentError are aliases in Py3)
        "FileNotFoundError" => &["FileNotFoundError", "OSError", "Exception", "BaseException"],
        "PermissionError" => &["PermissionError", "OSError", "Exception", "BaseException"],
        "TimeoutError" => &["TimeoutError", "OSError", "Exception", "BaseException"],
        "FileExistsError" => &["FileExistsError", "OSError", "Exception", "BaseException"],
        "IsADirectoryError" => &["IsADirectoryError", "OSError", "Exception", "BaseException"],
        "NotADirectoryError" => &[
            "NotADirectoryError",
            "OSError",
            "Exception",
            "BaseException",
        ],
        "BlockingIOError" => &["BlockingIOError", "OSError", "Exception", "BaseException"],
        "ConnectionError" => &["ConnectionError", "OSError", "Exception", "BaseException"],
        "BrokenPipeError" => &[
            "BrokenPipeError",
            "ConnectionError",
            "OSError",
            "Exception",
            "BaseException",
        ],
        "ConnectionResetError" => &[
            "ConnectionResetError",
            "ConnectionError",
            "OSError",
            "Exception",
            "BaseException",
        ],
        "ConnectionRefusedError" => &[
            "ConnectionRefusedError",
            "ConnectionError",
            "OSError",
            "Exception",
            "BaseException",
        ],
        "ConnectionAbortedError" => &[
            "ConnectionAbortedError",
            "ConnectionError",
            "OSError",
            "Exception",
            "BaseException",
        ],
        "ProcessLookupError" => &[
            "ProcessLookupError",
            "OSError",
            "Exception",
            "BaseException",
        ],
        "InterruptedError" => &["InterruptedError", "OSError", "Exception", "BaseException"],
        "ChildProcessError" => &["ChildProcessError", "OSError", "Exception", "BaseException"],
        "IOError" | "OSError" | "EnvironmentError" => &["OSError", "Exception", "BaseException"],
        // ValueError branch (Unicode* extend ValueError)
        "UnicodeDecodeError" => &[
            "UnicodeDecodeError",
            "UnicodeError",
            "ValueError",
            "Exception",
            "BaseException",
        ],
        "UnicodeEncodeError" => &[
            "UnicodeEncodeError",
            "UnicodeError",
            "ValueError",
            "Exception",
            "BaseException",
        ],
        "UnicodeError" => &["UnicodeError", "ValueError", "Exception", "BaseException"],
        "ValueError" => &["ValueError", "Exception", "BaseException"],
        // NameError / ImportError / SyntaxError branches
        "UnboundLocalError" => &[
            "UnboundLocalError",
            "NameError",
            "Exception",
            "BaseException",
        ],
        "NameError" => &["NameError", "Exception", "BaseException"],
        "ModuleNotFoundError" => &[
            "ModuleNotFoundError",
            "ImportError",
            "Exception",
            "BaseException",
        ],
        "ImportError" => &["ImportError", "Exception", "BaseException"],
        "IndentationError" => &[
            "IndentationError",
            "SyntaxError",
            "Exception",
            "BaseException",
        ],
        "TabError" => &[
            "TabError",
            "IndentationError",
            "SyntaxError",
            "Exception",
            "BaseException",
        ],
        "StopAsyncIteration" => &["StopAsyncIteration", "Exception", "BaseException"],
        _ => &[],
    }
}

/// Stamp `__types` with the canonical ancestor chain so typed catches
/// match base classes (`except LookupError:` catching a `KeyError`).
/// The multi-catch guard walks this array — same mechanism user classes
/// use via shared class emission. Unknown types get
/// `[canonical, "Exception", "BaseException"]`.
/// Stack: `[obj]` → `[obj]`
pub fn emit_stamp_exception_ancestors(chunk: &mut Chunk, exc_name: &str, line: u32) {
    let canon = canonical_exception_name(exc_name);
    let chain = exception_ancestors(exc_name);
    let fallback = [canon, "Exception", "BaseException"];
    let chain: &[&str] = if chain.is_empty() { &fallback } else { chain };
    chunk.emit_dup(line);
    for name in chain {
        chunk.emit_string_const(name, line);
    }
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, chain.len() as u16, line);
    let key = chunk.add_constant(Value::String(Arc::from(reflection::FIELD_TYPES)));
    chunk.emit_op_u16(Op::STRUCT_SET, key, line);
    chunk.emit_op(Op::DROP, line);
}

/// Emit the disposal half of a resource-management block (C# `using`,
/// Python `with`, Java try-with-resources, JS `using x = …`). Reads
/// the resource from `slot` and calls its lifecycle method (`Dispose`,
/// `__exit__`, `close`, …) if defined. Guards against the method
/// being absent so resources without a disposer don't trap.
///
/// ECMA-334 §13.14 / Python §8.5 / JS Stage 3 explicit-resource-
/// management share the same lowering: `try { body; } finally {
/// dispose; }`. We emit just the dispose tail; full try/finally
/// wrapping is the caller's job (or future enhancement here).
///
/// `dispose_method`: the canonical method name (`"Dispose"` for .NET,
/// `"__exit__"` for Python, `"close"` for Java AutoCloseable, etc.).
pub fn emit_resource_dispose(chunk: &mut Chunk, slot: u16, dispose_method: &str, line: u32) {
    let dispose_key = chunk.add_constant(Value::String(Arc::from(dispose_method)));
    let dispose_block = chunk.emit_block(line);
    // method = resource[<dispose_method>]
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, dispose_key, line);
    // if method is null/undefined, skip the call.
    chunk.emit_dup(line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_br_if(0, line);
    // Stack: [method]. Push receiver and CALL_REF(1). Drop result.
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_op_u8(Op::CALL_REF, 1, line);
    chunk.emit_op(Op::DROP, line);
    // Skipped path leaves `method` (null/undef) on stack — the END
    // closes the block, after which we DROP unconditionally.
    chunk.emit_end(line);
    chunk.patch_block(dispose_block);
    chunk.emit_op(Op::DROP, line);
}
