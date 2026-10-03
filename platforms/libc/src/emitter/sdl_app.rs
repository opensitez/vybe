//! SDL's synchronous entry runs on a WASM thread beside the browser UI loop.

use super::build::*;
use vybe_ast::{BinOp, ExprKind, ObjectProperty, Statement, StmtKind};

pub const WINDOWS: &str = "__libc_sdl_window_count";
pub const DOCUMENT: &str = "__libc_sdl_document";
pub const INPUT_QUEUE: &str = "__libc_sdl_input_queue";
pub const INPUT_STATE: &str = "__libc_sdl_input_state";

pub fn initialize_input(body: &mut Vec<Statement>) {
    body.insert(0, var_decl_stmt(DOCUMENT, int_lit(0)));
    body.insert(1, var_decl_stmt(INPUT_QUEUE, expr(ExprKind::Array(vec![]))));
    body.insert(
        2,
        var_decl_stmt(
            INPUT_STATE,
            expr(ExprKind::Object(
                [
                    "clientX", "clientY", "buttons", "shiftKey", "ctrlKey", "altKey", "metaKey",
                ]
                .into_iter()
                .map(|name| ObjectProperty::KeyValue {
                    key: str_lit(name),
                    value: int_lit(0),
                })
                .collect(),
            )),
        ),
    );
    let event_field = |name: &str| member(ident("event"), name);
    let state_field = |name: &str| member(ident(INPUT_STATE), name);
    let mouse_event = ["mousedown", "mouseup", "mousemove"]
        .into_iter()
        .map(|kind| binary(BinOp::Eq, event_field("type"), str_lit(kind)))
        .reduce(|left, right| binary(BinOp::Or, left, right))
        .unwrap();
    let mut input_body = vec![run(call_member(
        ident(INPUT_QUEUE),
        "push",
        vec![ident("event")],
    ))];
    input_body.push(if_stmt(
        mouse_event,
        ["clientX", "clientY", "buttons"]
            .into_iter()
            .map(|name| run(assign_expr(state_field(name), event_field(name))))
            .collect(),
        None,
    ));
    input_body.extend(
        ["shiftKey", "ctrlKey", "altKey", "metaKey"]
            .into_iter()
            .map(|name| run(assign_expr(state_field(name), event_field(name)))),
    );
    body.insert(
        3,
        function_stmt("__libc_sdl_input", vec!["event"], input_body),
    );
}

fn invoke(name: &str, args: Vec<vybe_ast::Expression>) -> vybe_ast::Expression {
    call_expr(ident(name), args)
}

fn run(value: vybe_ast::Expression) -> Statement {
    stmt(StmtKind::Expr(value))
}

fn set(name: &str, value: vybe_ast::Expression) -> Statement {
    run(assign_expr(ident(name), value))
}

fn binary(
    op: BinOp,
    left: vybe_ast::Expression,
    right: vybe_ast::Expression,
) -> vybe_ast::Expression {
    expr(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right),
    })
}

/// Run the SDL guest on the existing WASM thread path. The parent owns the
/// document and its event listeners; the child shares their queue and uses
/// the same document handle while Webcore keeps painting on the UI thread.
pub fn normalize_entry(body: &mut Vec<Statement>, entry_name: &str) {
    let Some(entry) = body.iter_mut().find(|s| {
        matches!(&s.kind,
        StmtKind::FunctionDecl { name, .. } if name == entry_name)
    }) else {
        return;
    };
    let StmtKind::FunctionDecl {
        body: guest_body, ..
    } = &mut entry.kind
    else {
        return;
    };
    let guest = function_stmt(
        "__sdl_guest",
        vec!["__sdl_thread_arg"],
        std::mem::take(guest_body),
    );
    *guest_body = vec![
        set(DOCUMENT, invoke("__c_sdl_document", vec![])),
        run(invoke(
            "__c_sdl_listen_input",
            vec![ident("__libc_sdl_input")],
        )),
        run(invoke(
            "__c_thread_spawn",
            vec![int_lit(0), expr(ExprKind::FunctionExpr(Box::new(guest)))],
        )),
        stmt(StmtKind::Return(Some(int_lit(0)))),
    ];
    body.insert(0, var_decl_stmt(WINDOWS, int_lit(0)));
}

pub fn emit_window_count(
    chunks: &mut [vybe_runtime::Chunk],
    current: usize,
    delta: i32,
    line: u32,
) {
    use vybe_compiler::primitives::globals;
    use vybe_runtime::Op;
    let chunk = &mut chunks[current];
    globals::emit_read(chunk, WINDOWS, line);
    chunk.emit_f64_const(delta as f64, line);
    chunk.emit_op(Op::F64_ADD, line);
    globals::emit_write(chunk, WINDOWS, line);
}

pub fn emit_listen_input(chunks: &mut [vybe_runtime::Chunk], current: usize, line: u32) {
    use vybe_runtime::Op;
    let chunk = &mut chunks[current];
    let callback = chunk.alloc_scratch(1);
    let document_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, callback, line);
    let document = chunk.add_import("web:html", "activeDocument");
    let listen = chunk.add_import("web:dom", "addEventListener");
    chunk.emit_call(document, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, document_slot, line);
    for kind in [
        "keydown",
        "keyup",
        "mousedown",
        "mouseup",
        "mousemove",
        "wheel",
    ] {
        chunk.emit_op_u16(Op::LOCAL_GET, document_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, document_slot, line);
        chunk.emit_string_const(kind, line);
        chunk.emit_op_u16(Op::LOCAL_GET, callback, line);
        chunk.emit_call(listen, 4, line);
    }
    chunk.emit_i32_const(0, line);
}
