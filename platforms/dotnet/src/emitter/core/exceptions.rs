//! The .NET exception hierarchy as tree adapter classes.
//!
//! Every row of [`EXCEPTION_HIERARCHY`] is a `dotnet.System` component class
//! whose constructor is backed by [`emit_exception_ctor`]. The instance is an
//! object of its reserved WASM type — the hierarchy is the supertype chain, so
//! a typed catch, `TypeOf … Is` and `GetType()` are `ref.test` — built by the
//! shared exception machinery in `errors.rs`. This adapter adds only what .NET
//! surfaces on top: the constructor argument orders, `ParamName` /
//! `ActualValue` / `ObjectName`, and the message forms those constructors
//! compose.

use vybe_compiler::primitives::class_slots::{self, Dest, ObjSource, ValueSource};
use vybe_compiler::primitives::errors;
use vybe_compiler::primitives::instructions::host;
use vybe_compiler::primitives::ops;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

use crate::emitter::core::object_fields::{field_slot, set_both_spellings};

pub const EXCEPTION_HIERARCHY: &[(&str, &str)] = &[
    ("Exception", ""),
    ("SystemException", "Exception"),
    ("ApplicationException", "Exception"),
    ("TargetInvocationException", "ApplicationException"),
    ("AggregateException", "Exception"),
    ("ArithmeticException", "SystemException"),
    ("ArgumentException", "SystemException"),
    ("InvalidOperationException", "SystemException"),
    ("SecurityException", "SystemException"),
    ("TypeInitializationException", "SystemException"),
    ("FormatException", "SystemException"),
    ("InvalidCastException", "SystemException"),
    ("NotImplementedException", "SystemException"),
    ("NotSupportedException", "SystemException"),
    ("ParameterBindingException", "SystemException"),
    ("NullReferenceException", "SystemException"),
    ("IndexOutOfRangeException", "SystemException"),
    ("KeyNotFoundException", "SystemException"),
    ("IOException", "SystemException"),
    ("ItemNotFoundException", "SystemException"),
    ("TimeoutException", "SystemException"),
    ("OperationCanceledException", "SystemException"),
    ("DivideByZeroException", "ArithmeticException"),
    ("OverflowException", "ArithmeticException"),
    ("ArgumentNullException", "ArgumentException"),
    ("ArgumentOutOfRangeException", "ArgumentException"),
    ("FileNotFoundException", "IOException"),
    ("DirectoryNotFoundException", "IOException"),
    ("ObjectDisposedException", "InvalidOperationException"),
    ("UriFormatException", "FormatException"),
    ("XmlException", "SystemException"),
];

/// Prefix of the `Common` constructor backing each hierarchy row declares;
/// the class name follows it.
pub const EXCEPTION_CTOR_PREFIX: &str = "dotnet.exception_new.";

pub fn exception_ctor_backing(name: &str) -> String {
    format!("{EXCEPTION_CTOR_PREFIX}{name}")
}

/// Whether `name` is a row of the hierarchy — a `dotnet.System` tree class
/// reachable under its short name from every .NET language.
pub fn is_synthesized_exception_class(name: &str) -> bool {
    EXCEPTION_HIERARCHY
        .iter()
        .any(|(class, _)| class.eq_ignore_ascii_case(name))
}

/// The class and its ancestors, leaf first — the supertype chain of its
/// WASM type. A name outside the table is a direct child of `Exception`.
pub fn exception_type_chain(name: &str) -> Vec<String> {
    let mut chain = vec![name.to_string()];
    let mut current = name;
    loop {
        let parent = EXCEPTION_HIERARCHY
            .iter()
            .find_map(|(child, parent)| child.eq_ignore_ascii_case(current).then_some(*parent));
        match parent {
            Some("") => break,
            Some(parent) => {
                chain.push(parent.to_string());
                current = parent;
            }
            None => {
                if !name.eq_ignore_ascii_case("Exception") {
                    chain.push("Exception".to_string());
                }
                break;
            }
        }
    }
    chain
}

/// A typed instance of `name` with `msg` as its message, on the stack —
/// the one allocation every dotnet adapter uses to raise.
pub fn emit_new_typed(
    chunks: &mut [Chunk],
    current: usize,
    name: &str,
    msg: ValueSource,
    line: u32,
) {
    let chain = exception_type_chain(name);
    errors::emit_exception_new_typed(chunks, current, name, &chain, msg, line);
}

/// Raises a typed `name` with a constant message.
pub fn emit_throw_typed(
    chunks: &mut [Chunk],
    current: usize,
    name: &str,
    message: &str,
    line: u32,
) {
    emit_new_typed(
        chunks,
        current,
        name,
        ValueSource::ConstStr(message.to_string()),
        line,
    );
    errors::emit_throw(&mut chunks[current], line);
}

/// Where the constructed instance comes from.
#[derive(Clone, Copy)]
enum Target {
    /// A fresh object of the class's type.
    New,
    /// An object a derived constructor already allocated, in this local.
    Into(u16),
}

/// `new <class>(args…)` with `argc` arguments on the stack, in order.
/// Leaves the instance on the stack.
pub fn emit_exception_ctor(chunks: &mut [Chunk], current: usize, class: &str, argc: u8, line: u32) {
    emit_ctor(chunks, current, class, argc, Target::New, line);
}

/// The `#into` form: the receiver a derived constructor allocated is on top
/// of the stack above the `argc` arguments; it is initialised as `<class>`
/// and left on the stack.
pub fn emit_exception_ctor_into(
    chunks: &mut [Chunk],
    current: usize,
    class: &str,
    argc: u8,
    line: u32,
) {
    let target = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, target, line);
    emit_ctor(chunks, current, class, argc, Target::Into(target), line);
}

fn emit_ctor(
    chunks: &mut [Chunk],
    current: usize,
    class: &str,
    argc: u8,
    target: Target,
    line: u32,
) {
    let args: Vec<u16> = (0..argc)
        .map(|_| chunks[current].alloc_scratch(1))
        .collect();
    for slot in args.iter().rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, *slot, line);
    }
    let exc = chunks[current].alloc_scratch(1);
    let arg = |index: usize| args.get(index).copied();
    let c = &mut chunks[current];

    match (class, arg(0), arg(1), arg(2)) {
        // ArgumentNullException(paramName) · (paramName, message) · (message, inner)
        ("ArgumentNullException", Some(param), None, _) => {
            emit_message_with_param(c, Message::Const("Value cannot be null."), param, line);
            emit_new(chunks, current, class, target, exc, line);
            set_both_spellings(&mut chunks[current], exc, param, "ParamName", line);
        }
        ("ArgumentNullException", Some(first), Some(second), _) => {
            emit_is_string(c, second, line);
            c.emit_if(line);
            emit_message_with_param(c, Message::Local(second), first, line);
            emit_new(chunks, current, class, target, exc, line);
            set_both_spellings(&mut chunks[current], exc, first, "ParamName", line);
            chunks[current].emit_else(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, first, line);
            emit_new(chunks, current, class, target, exc, line);
            emit_attach_inner(&mut chunks[current], exc, second, line);
            chunks[current].emit_end(line);
        }
        // ArgumentOutOfRangeException(paramName) · (paramName, message) ·
        // (message, inner) · (paramName, actualValue, message)
        ("ArgumentOutOfRangeException", Some(param), None, _) => {
            emit_message_with_param(
                c,
                Message::Const("Specified argument was out of the range of valid values."),
                param,
                line,
            );
            emit_new(chunks, current, class, target, exc, line);
            set_both_spellings(&mut chunks[current], exc, param, "ParamName", line);
        }
        ("ArgumentOutOfRangeException", Some(first), Some(second), None) => {
            emit_is_string(c, second, line);
            c.emit_if(line);
            emit_message_with_param(c, Message::Local(second), first, line);
            emit_new(chunks, current, class, target, exc, line);
            set_both_spellings(&mut chunks[current], exc, first, "ParamName", line);
            chunks[current].emit_else(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, first, line);
            emit_new(chunks, current, class, target, exc, line);
            emit_attach_inner(&mut chunks[current], exc, second, line);
            chunks[current].emit_end(line);
        }
        ("ArgumentOutOfRangeException", Some(param), Some(actual), Some(message)) => {
            emit_message_with_param(c, Message::Local(message), param, line);
            c.emit_string_const("\nActual value was ", line);
            ops::emit_dyn_add(c, line);
            c.emit_op_u16(Op::LOCAL_GET, actual, line);
            let to_string = c.add_import("ecma:string", "String");
            c.emit_call(to_string, 1, line);
            ops::emit_dyn_add(c, line);
            c.emit_string_const(".", line);
            ops::emit_dyn_add(c, line);
            emit_new(chunks, current, class, target, exc, line);
            set_both_spellings(&mut chunks[current], exc, param, "ParamName", line);
            set_both_spellings(&mut chunks[current], exc, actual, "ActualValue", line);
        }
        // ArgumentException(message) · (message, inner) · (message, paramName) ·
        // (message, paramName, inner)
        ("ArgumentException", Some(message), Some(second), inner) => {
            emit_is_string(c, second, line);
            c.emit_if(line);
            emit_message_with_param(c, Message::Local(message), second, line);
            emit_new(chunks, current, class, target, exc, line);
            set_both_spellings(&mut chunks[current], exc, second, "ParamName", line);
            if let Some(inner) = inner {
                emit_attach_inner(&mut chunks[current], exc, inner, line);
            }
            chunks[current].emit_else(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, message, line);
            emit_new(chunks, current, class, target, exc, line);
            emit_attach_inner(&mut chunks[current], exc, second, line);
            chunks[current].emit_end(line);
        }
        // ObjectDisposedException(objectName) · (objectName, message) · (message, inner)
        ("ObjectDisposedException", Some(object_name), None, _) => {
            c.emit_string_const("Cannot access a disposed object.\nObject name: '", line);
            c.emit_op_u16(Op::LOCAL_GET, object_name, line);
            ops::emit_dyn_add(c, line);
            c.emit_string_const("'.", line);
            ops::emit_dyn_add(c, line);
            emit_new(chunks, current, class, target, exc, line);
            set_both_spellings(&mut chunks[current], exc, object_name, "ObjectName", line);
        }
        ("ObjectDisposedException", Some(first), Some(second), _) => {
            emit_is_string(c, second, line);
            c.emit_if(line);
            c.emit_op_u16(Op::LOCAL_GET, second, line);
            emit_new(chunks, current, class, target, exc, line);
            set_both_spellings(&mut chunks[current], exc, first, "ObjectName", line);
            chunks[current].emit_else(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, first, line);
            emit_new(chunks, current, class, target, exc, line);
            emit_attach_inner(&mut chunks[current], exc, second, line);
            chunks[current].emit_end(line);
        }
        // Every other row: (message) · (message, inner)
        (_, message, inner, _) => {
            match message {
                Some(message) => c.emit_op_u16(Op::LOCAL_GET, message, line),
                None => c.emit_string_const("", line),
            }
            emit_new(chunks, current, class, target, exc, line);
            if let Some(inner) = inner {
                emit_attach_inner(&mut chunks[current], exc, inner, line);
            }
        }
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, exc, line);
}

enum Message<'a> {
    Const(&'a str),
    Local(u16),
}

/// Pushes `<message> (Parameter '<param>')`, the form .NET composes for the
/// argument exceptions.
fn emit_message_with_param(chunk: &mut Chunk, message: Message<'_>, param: u16, line: u32) {
    match message {
        Message::Const(text) => chunk.emit_string_const(text, line),
        Message::Local(slot) => chunk.emit_op_u16(Op::LOCAL_GET, slot, line),
    }
    chunk.emit_string_const(" (Parameter '", line);
    ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, param, line);
    ops::emit_dyn_add(chunk, line);
    chunk.emit_string_const("')", line);
    ops::emit_dyn_add(chunk, line);
}

/// Message on the stack → the instance in `exc`, through the shared machinery:
/// a fresh object of the class's reserved type, or the derived receiver.
fn emit_new(
    chunks: &mut [Chunk],
    current: usize,
    class: &str,
    target: Target,
    exc: u16,
    line: u32,
) {
    match target {
        Target::New => {
            let chain = exception_type_chain(class);
            errors::emit_exception_new_typed(
                chunks,
                current,
                class,
                &chain,
                ValueSource::Stack,
                line,
            );
        }
        Target::Into(receiver) => {
            let c = &mut chunks[current];
            let message = c.alloc_scratch(1);
            c.emit_op_u16(Op::LOCAL_SET, message, line);
            c.emit_op_u16(Op::LOCAL_GET, receiver, line);
            c.emit_dup(line);
            c.emit_op_u16(Op::LOCAL_GET, message, line);
            errors::emit_exception_new_finalize(c, class, line);
        }
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, exc, line);
}

/// The shared attach writes `cause` and `InnerException`; a case-insensitive
/// language reads the folded spelling, so that one is written too.
fn emit_attach_inner(chunk: &mut Chunk, exc: u16, inner: u16, line: u32) {
    errors::emit_attach_second_ctor_arg(chunk, exc, inner, line);
    chunk.emit_op_u16(Op::LOCAL_GET, exc, line);
    chunk.emit_op_u16(Op::LOCAL_GET, exc, line);
    class_slots::emit_class_get(
        chunk,
        ObjSource::Stack,
        &field_slot("InnerException"),
        Dest::Stack,
        line,
    );
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &field_slot("innerexception"),
        ValueSource::Stack,
        line,
    );
}

fn emit_is_string(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    host::emit(chunk, "ecma:value", "typeof", 1, line);
    chunk.emit_string_const("string", line);
    ops::emit_dyn_eq(chunk, line);
}
