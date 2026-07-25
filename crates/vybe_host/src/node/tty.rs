//! `node:tty` — Node.js TTY module.
//!
//! Reference: <https://nodejs.org/api/tty.html>.

use vybe_bytecode::VM;
use vybe_bytecode::value::{Object, Value};

#[cfg(unix)]
fn is_tty_fd(fd: i32) -> bool {
    unsafe extern "C" {
        fn isatty(fd: i32) -> i32;
    }
    unsafe { isatty(fd) != 0 }
}

#[cfg(not(unix))]
fn is_tty_fd(_fd: i32) -> bool {
    false
}

fn make_read_stream(fd: i32) -> Value {
    let mut o = Object::new();
    o.properties.insert("fd".into(), Value::I32(fd));
    o.properties
        .insert("isTTY".into(), Value::Bool(is_tty_fd(fd)));
    o.properties.insert("isRaw".into(), Value::Bool(false));
    for m in [
        "setRawMode",
        "on",
        "once",
        "off",
        "destroy",
        "pause",
        "resume",
        "emit",
        "addListener",
        "removeListener",
    ] {
        o.properties.insert(m.into(), Value::Undefined);
    }
    Value::Object(vybe_bytecode::heap::alloc(o))
}

fn make_write_stream(fd: i32) -> Value {
    let mut o = Object::new();
    o.properties.insert("fd".into(), Value::I32(fd));
    o.properties
        .insert("isTTY".into(), Value::Bool(is_tty_fd(fd)));
    o.properties.insert("columns".into(), Value::I32(80));
    o.properties.insert("rows".into(), Value::I32(24));
    for m in [
        "clearLine",
        "cursorTo",
        "moveCursor",
        "getColorDepth",
        "hasColors",
        "on",
        "once",
        "off",
        "emit",
        "write",
        "destroy",
        "addListener",
        "removeListener",
    ] {
        o.properties.insert(m.into(), Value::Undefined);
    }
    Value::Object(vybe_bytecode::heap::alloc(o))
}

pub fn register(vm: &mut VM) {
    vm.register_host_fn(
        "node:tty",
        "isatty",
        Box::new(|_ctx, args| match args.first() {
            Some(Value::F64(f)) if f.fract() != 0.0 => Value::Bool(false),
            Some(Value::I32(fd)) => Value::Bool(is_tty_fd(*fd)),
            Some(Value::F64(f)) => {
                let fd = *f as i32;
                if fd < 0 {
                    Value::Bool(false)
                } else {
                    Value::Bool(is_tty_fd(fd))
                }
            }
            _ => Value::Bool(false),
        }),
    );

    vm.register_host_fn(
        "node:tty",
        "ReadStream",
        Box::new(|_ctx, args| {
            let fd = match args.first() {
                Some(Value::I32(n)) => *n,
                _ => 0,
            };
            make_read_stream(fd)
        }),
    );

    vm.register_host_fn(
        "node:tty",
        "WriteStream",
        Box::new(|_ctx, args| {
            let fd = match args.first() {
                Some(Value::I32(n)) => *n,
                _ => 1,
            };
            make_write_stream(fd)
        }),
    );

    vm.register_host_fn(
        "node:tty",
        "getColorDepth",
        Box::new(|_ctx, _args| Value::I32(8)),
    );

    vm.register_host_fn(
        "node:tty",
        "hasColors",
        Box::new(|_ctx, args| {
            let count = match args.first() {
                Some(Value::I32(n)) => *n,
                Some(Value::F64(f)) => *f as i32,
                _ => 2,
            };
            Value::Bool(count <= 256)
        }),
    );
}
