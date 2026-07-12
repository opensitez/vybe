//! Event-handler AST lowering — language-agnostic helpers that normalise
//! event-subscription syntax (`AddHandler`, `control.Click += handler`, …) onto
//! the canonical `StmtKind::AddHandler`/`RemoveHandler` nodes. Kept separate
//! from `gui.rs` because events are not necessarily a GUI concern.
//!
//! The emit side lives in the compiler (`compiler/events.rs`) and the shared
//! `vybe:gui` binding in `gui.rs`; these are the *walker-facing* builders any
//! front-end can reuse.

use vybe_ast::{BinOp, ExprKind, Expression, StmtKind};

pub fn add_handler_stmt(
    control: Expression,
    event: impl Into<String>,
    handler: Expression,
) -> StmtKind {
    StmtKind::AddHandler {
        control,
        event: event.into(),
        handler,
    }
}

pub fn remove_handler_stmt(
    control: Expression,
    event: impl Into<String>,
    handler: Expression,
) -> StmtKind {
    StmtKind::RemoveHandler {
        control,
        event: event.into(),
        handler,
    }
}

pub fn lower_event_compound_assignment(expr: &Expression) -> Option<StmtKind> {
    let ExprKind::Assign { target, value } = &expr.kind else {
        return None;
    };
    let ExprKind::Member {
        object: ev_obj,
        field: ev_field,
        ..
    } = &target.kind
    else {
        return None;
    };
    let ExprKind::Binary { op, left, right } = &value.kind else {
        return None;
    };

    let same_target = matches!(
        &left.kind,
        ExprKind::Member { object, field, .. } if member_eq(object, field, ev_obj, ev_field)
    );
    let handler = unwrap_event_handler(right)?;

    if !same_target {
        return None;
    }

    let event_name = ev_field.to_lowercase();
    let control = (**ev_obj).clone();
    Some(match op {
        BinOp::Add => add_handler_stmt(control, event_name, handler.clone()),
        BinOp::Sub => remove_handler_stmt(control, event_name, handler.clone()),
        _ => return None,
    })
}

fn unwrap_event_handler(expr: &Expression) -> Option<&Expression> {
    if is_event_handler_expr(expr) {
        return Some(expr);
    }

    match &expr.kind {
        ExprKind::New { args, .. } if args.len() == 1 => {
            let inner = &args[0].value;
            if is_event_handler_expr(inner) {
                Some(inner)
            } else {
                None
            }
        }
        _ => None,
    }
}

pub fn is_event_handler_expr(expr: &Expression) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Ident(_) | ExprKind::Member { .. } | ExprKind::Lambda { .. }
    )
}

/// True if `field` is a known WinForms/GUI event name (case-insensitive).
/// Used by language walkers to recognise `control.Click += handler` as event
/// subscription rather than ordinary numeric compound-assignment.
pub fn is_known_gui_event_field(field: &str) -> bool {
    matches!(
        field.to_lowercase().as_str(),
        "click"
            | "dblclick"
            | "doubleclick"
            | "load"
            | "unload"
            | "change"
            | "textchanged"
            | "selectedindexchanged"
            | "checkedchanged"
            | "valuechanged"
            | "keypress"
            | "keydown"
            | "keyup"
            | "mousedown"
            | "mouseup"
            | "mousemove"
            | "mouseenter"
            | "mouseleave"
            | "gotfocus"
            | "lostfocus"
            | "enter"
            | "leave"
            | "resize"
            | "paint"
            | "formclosing"
    )
}

fn member_eq(obj_a: &Expression, field_a: &str, obj_b: &Expression, field_b: &str) -> bool {
    if !field_a.eq_ignore_ascii_case(field_b) {
        return false;
    }

    match (&obj_a.kind, &obj_b.kind) {
        (ExprKind::Ident(a), ExprKind::Ident(b)) => a == b,
        (ExprKind::This, ExprKind::This) => true,
        (
            ExprKind::Member {
                object: inner_a,
                field: inner_field_a,
                ..
            },
            ExprKind::Member {
                object: inner_b,
                field: inner_field_b,
                ..
            },
        ) => member_eq(inner_a, inner_field_a, inner_b, inner_field_b),
        _ => false,
    }
}
