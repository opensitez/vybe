//! `sched` — event scheduler backed by normal Python class/collection shapes.

use super::builders::*;
use vybe_ast::{Argument, BinOp, ExprKind, Expression, Statement};

type Expr = Expression;

fn op(o: BinOp, left: Expr, right: Expr) -> Expr {
    binary(o, left, right)
}

fn and(left: Expr, right: Expr) -> Expr {
    op(BinOp::And, left, right)
}

fn or(left: Expr, right: Expr) -> Expr {
    op(BinOp::Or, left, right)
}

fn self_attr(name: &str) -> Expr {
    this_field(name)
}

fn event_time(event: Expr) -> Expr {
    field_of(event, "time")
}

fn event_priority(event: Expr) -> Expr {
    field_of(event, "priority")
}

fn event_before(left: Expr, right: Expr) -> Expr {
    or(
        op(BinOp::Lt, event_time(left.clone()), event_time(right.clone())),
        and(
            op(BinOp::Eq, event_time(left.clone()), event_time(right.clone())),
            op(BinOp::Lt, event_priority(left), event_priority(right)),
        ),
    )
}

fn event_compare(opcode: BinOp) -> Vec<Statement> {
    let self_event = ident("self");
    let other_event = ident("other");
    let time_cmp = op(
        opcode,
        event_time(self_event.clone()),
        event_time(other_event.clone()),
    );
    let priority_cmp = op(
        opcode,
        event_priority(self_event.clone()),
        event_priority(other_event.clone()),
    );

    match opcode {
        BinOp::Lt | BinOp::Gt => vec![ret(or(
            time_cmp,
            and(
                op(BinOp::Eq, event_time(self_event), event_time(other_event)),
                priority_cmp,
            ),
        ))],
        BinOp::LtEq | BinOp::GtEq => vec![ret(or(
            time_cmp,
            and(
                op(BinOp::Eq, event_time(self_event.clone()), event_time(other_event.clone())),
                op(
                    opcode,
                    event_priority(self_event),
                    event_priority(other_event),
                ),
            ),
        ))],
        BinOp::Eq => vec![ret(and(
            op(BinOp::Eq, event_time(self_event), event_time(other_event)),
            op(BinOp::Eq, event_priority(ident("self")), event_priority(ident("other"))),
        ))],
        _ => vec![ret(bool_lit(false))],
    }
}

fn append_to(queue: Expr, event: Expr) -> Statement {
    expr_stmt(call(member(queue, "append"), vec![event]))
}

fn sorted_insert_event(event: Expr) -> Vec<Statement> {
    vec![
        assign(ident("__new_queue"), list_of(vec![])),
        assign(ident("__inserted"), bool_lit(false)),
        for_in(
            "__queued_event",
            self_attr("_queue"),
            vec![
                if_stmt(
                    and(
                        unary_not(ident("__inserted")),
                        event_before(event.clone(), ident("__queued_event")),
                    ),
                    vec![
                        append_to(ident("__new_queue"), event.clone()),
                        assign(ident("__inserted"), bool_lit(true)),
                    ],
                ),
                append_to(ident("__new_queue"), ident("__queued_event")),
            ],
        ),
        if_stmt(
            unary_not(ident("__inserted")),
            vec![append_to(ident("__new_queue"), event)],
        ),
        set_this("_queue", ident("__new_queue")),
        refresh_queue_view(),
    ]
}

fn cancel_event(event: Expr) -> Vec<Statement> {
    vec![
        assign(ident("__new_queue"), list_of(vec![])),
        assign(ident("__found"), bool_lit(false)),
        for_in(
            "__queued_event",
            self_attr("_queue"),
            vec![
                if_stmt(
                    call_global("__py_is__", vec![ident("__queued_event"), event.clone()]),
                    vec![assign(ident("__found"), bool_lit(true))],
                ),
                if_stmt(
                    call_global("__py_is_not__", vec![ident("__queued_event"), event.clone()]),
                    vec![append_to(ident("__new_queue"), ident("__queued_event"))],
                ),
            ],
        ),
        if_stmt(
            unary_not(ident("__found")),
            vec![raise_call("ValueError", vec![])],
        ),
        set_this("_queue", ident("__new_queue")),
        refresh_queue_view(),
        ret(null()),
    ]
}

fn queue_len() -> Expr {
    call_global("len", vec![self_attr("_queue")])
}

fn first_event() -> Expr {
    index(self_attr("_queue"), num(0.0))
}

fn refresh_queue_view() -> Statement {
    set_this("queue", call_global("list", vec![self_attr("_queue")]))
}

fn default_timefunc_lambda() -> Expr {
    Expression::with_span(
        ExprKind::Lambda {
            params: vec![],
            body: vybe_ast::LambdaBody::Expr(Box::new(num(0.0))),
            is_async: false,
            captures: vec![],
        },
        span(),
    )
}

fn default_delayfunc_lambda() -> Expr {
    Expression::with_span(
        ExprKind::Lambda {
            params: vec![param("delay", None)],
            body: vybe_ast::LambdaBody::Expr(Box::new(null())),
            is_async: false,
            captures: vec![],
        },
        span(),
    )
}

fn execute_event(event: Expr) -> Statement {
    expr_stmt(call_global(
        "__py_sched_execute",
        vec![
            field_of(event.clone(), "action"),
            field_of(event.clone(), "argument"),
            field_of(event, "kwargs"),
        ],
    ))
}

fn named_arg(name: &str, value: Expr) -> Argument {
    Argument {
        value,
        name: Some(name.to_string()),
        by_ref: false,
        spread: false,
    }
}

fn named_call(callee: Expr, args: Vec<Argument>) -> Expr {
    Expression::with_span(
        ExprKind::Call {
            callee: Box::new(callee),
            args,
            optional: false,
        },
        span(),
    )
}

fn dict_get(dict: Expr, key: &str, default: Expr) -> Expr {
    call(member(dict, "get"), vec![str_lit(key), default])
}

fn sched_apply(action: Expr, argument: Expr) -> Expr {
    call_global("__py_reflect_apply", vec![action, null(), argument])
}

pub(super) fn event() -> Statement {
    class(
        "SchedEvent",
        vec![
            init(
                vec![
                    param("time", None),
                    param("priority", None),
                    param("action", None),
                    param("argument", None),
                    param("kwargs", None),
                ],
                vec![
                    set_this("time", ident("time")),
                    set_this("priority", ident("priority")),
                    set_this("action", ident("action")),
                    set_this("argument", ident("argument")),
                    set_this("kwargs", ident("kwargs")),
                ],
            ),
            method("__lt__", vec![param("other", None)], event_compare(BinOp::Lt)),
            method("__le__", vec![param("other", None)], event_compare(BinOp::LtEq)),
            method("__gt__", vec![param("other", None)], event_compare(BinOp::Gt)),
            method("__ge__", vec![param("other", None)], event_compare(BinOp::GtEq)),
            method("__eq__", vec![param("other", None)], event_compare(BinOp::Eq)),
        ],
    )
}

pub(super) fn scheduler() -> Statement {
    let mut enterabs_body = vec![
        if_stmt(
            is_none(ident("argument")),
            vec![assign(ident("argument"), tuple_of(vec![]))],
        ),
        if_stmt(
            is_none(ident("kwargs")),
            vec![assign(ident("kwargs"), dict_str(vec![]))],
        ),
        assign(
            ident("__event"),
            new(
                "SchedEvent",
                vec![
                    ident("time"),
                    ident("priority"),
                    ident("action"),
                    ident("argument"),
                    ident("kwargs"),
                ],
            ),
        ),
    ];
    enterabs_body.extend(sorted_insert_event(ident("__event")));
    enterabs_body.push(ret(ident("__event")));

    class(
        "scheduler",
        vec![
            init(
                vec![
                    param("timefunc", Some(null())),
                    param("delayfunc", Some(null())),
                ],
                vec![
                    set_this(
                        "timefunc",
                        ternary(
                            is_none(ident("timefunc")),
                            default_timefunc_lambda(),
                            ident("timefunc"),
                        ),
                    ),
                    set_this(
                        "delayfunc",
                        ternary(
                            is_none(ident("delayfunc")),
                            default_delayfunc_lambda(),
                            ident("delayfunc"),
                        ),
                    ),
                    set_this("_queue", list_of(vec![])),
                    refresh_queue_view(),
                ],
            ),
            method(
                "enterabs",
                vec![
                    param("time", None),
                    param("priority", None),
                    param("action", None),
                    param("argument", Some(null())),
                    param("kwargs", Some(null())),
                ],
                enterabs_body,
            ),
            method(
                "enter",
                vec![
                    param("delay", None),
                    param("priority", None),
                    param("action", None),
                    param("argument", Some(null())),
                    param("kwargs", Some(null())),
                ],
                vec![ret(call(
                    member(ident("self"), "enterabs"),
                    vec![
                        op(BinOp::Add, call(self_attr("timefunc"), vec![]), ident("delay")),
                        ident("priority"),
                        ident("action"),
                        ident("argument"),
                        ident("kwargs"),
                    ],
                ))],
            ),
            method(
                "cancel",
                vec![param("event", None)],
                cancel_event(ident("event")),
            ),
            method(
                "empty",
                vec![],
                vec![ret(op(BinOp::Eq, queue_len(), num(0.0)))],
            ),
            method(
                "run",
                vec![param("blocking", Some(bool_lit(true)))],
                vec![
                    while_stmt(
                        op(BinOp::Gt, queue_len(), num(0.0)),
                        vec![
                            assign(ident("__event"), first_event()),
                            assign(ident("__now"), call(self_attr("timefunc"), vec![])),
                            if_stmt(
                                and(
                                    unary_not(ident("blocking")),
                                    op(BinOp::Gt, event_time(ident("__event")), ident("__now")),
                                ),
                                vec![ret(null())],
                            ),
                            if_stmt(
                                op(BinOp::Gt, event_time(ident("__event")), ident("__now")),
                                vec![expr_stmt(call(
                                    self_attr("delayfunc"),
                                    vec![op(
                                        BinOp::Sub,
                                        event_time(ident("__event")),
                                        ident("__now"),
                                    )],
                                ))],
                            ),
                            set_this("_queue", slice_from(self_attr("_queue"), num(1.0))),
                            refresh_queue_view(),
                            execute_event(ident("__event")),
                            expr_stmt(call(self_attr("delayfunc"), vec![num(0.0)])),
                        ],
                    ),
                    ret(null()),
                ],
            ),
        ],
    )
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![function(
        "__py_sched_execute",
        vec![
            param("action", None),
            param("argument", None),
            param("kwargs", None),
        ],
        vec![
            if_stmt(
                op(BinOp::Eq, call_global("len", vec![ident("kwargs")]), num(0.0)),
                vec![ret(sched_apply(ident("action"), ident("argument")))],
            ),
            if_stmt(
                and(
                    is_not_none(dict_get(ident("kwargs"), "name", null())),
                    is_not_none(dict_get(ident("kwargs"), "val", null())),
                ),
                vec![ret(named_call(
                    ident("action"),
                    vec![
                        named_arg("name", dict_get(ident("kwargs"), "name", null())),
                        named_arg("val", dict_get(ident("kwargs"), "val", null())),
                    ],
                ))],
            ),
            ret(sched_apply(ident("action"), ident("argument"))),
        ],
    )]
}
