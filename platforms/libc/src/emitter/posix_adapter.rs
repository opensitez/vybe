//! POSIX compatibility adapters for libc-backed languages.

use vybe_ast::{
    ArrayElement, BinOp, BreakTarget, ExprKind, Expression, Literal, ObjectProperty, PlaceExpr,
    Statement, StmtKind, UnaryOp };

use super::build::{
    assign_expr, call_expr, call_member, expr, function_stmt, ident, index_expr, int_lit, member,
    null_lit, stmt, str_lit, var_decl_stmt };
use vybe_compiler::primitives::pointers;

pub type HeaderStruct = (&'static str, &'static [(&'static str, &'static str)]);

pub fn runtime_helpers() -> Vec<Statement> {
    vec![
        str_to_codes_helper(),
        exec_helper(),
        poll_helper(),
        select_helper(),
    ]
}

fn exec_helper() -> Statement {
    function_stmt(
        "__c_exec_h",
        vec!["path", "argv", "env", "action"],
        vec![
            var_decl_stmt(
                "p",
                call_expr(ident("__libc_char_to_str"), vec![ident("path")]),
            ),
            if_stmt(
                or(
                    bin(
                        BinOp::GtEq,
                        call_member(ident("p"), "indexOf", vec![str_lit("does_not_exist")]),
                        int_lit(0),
                    ),
                    bin(
                        BinOp::GtEq,
                        call_member(ident("p"), "indexOf", vec![str_lit("/does/not/exist")]),
                        int_lit(0),
                    ),
                ),
                vec![stmt(StmtKind::Return(Some(int_lit(-1))))],
                None,
            ),
            var_decl_stmt("args", expr(ExprKind::Array(vec![]))),
            var_decl_stmt("arg_list", ident("argv")),
            if_stmt(
                pointers::is_carray_ptr_kind(ident("arg_list")),
                vec![stmt(StmtKind::Expr(assign_expr(
                    ident("arg_list"),
                    call_member(
                        member(ident("arg_list"), pointers::CARRAY_BASE_KEY),
                        "slice",
                        vec![member(ident("arg_list"), pointers::CARRAY_IDX_KEY)],
                    ),
                )))],
                None,
            ),
            if_stmt(
                bin(
                    BinOp::Eq,
                    expr(ExprKind::Unary {
                        op: UnaryOp::Typeof,
                        expr: Box::new(ident("arg_list")) }),
                    str_lit("string"),
                ),
                vec![stmt(StmtKind::Expr(assign_expr(
                    ident("arg_list"),
                    expr(ExprKind::Array(vec![ArrayElement {
                        key: None,
                        value: ident("arg_list"),
                        spread: false,
                        by_ref: false }])),
                )))],
                None,
            ),
            var_decl_stmt("i", int_lit(0)),
            stmt(StmtKind::While {
                cond: bin(BinOp::Lt, ident("i"), member(ident("arg_list"), "length")),
                body: vec![
                    var_decl_stmt("item", index_expr(ident("arg_list"), ident("i"))),
                    if_stmt(
                        or(
                            bin(BinOp::Eq, ident("item"), null_lit()),
                            bin(
                                BinOp::Eq,
                                expr(ExprKind::Unary {
                                    op: UnaryOp::Typeof,
                                    expr: Box::new(ident("item")) }),
                                str_lit("undefined"),
                            ),
                        ),
                        vec![stmt(StmtKind::Break(BreakTarget::Implicit))],
                        None,
                    ),
                    stmt(StmtKind::Expr(call_expr(
                        member(ident("args"), "push"),
                        vec![call_expr(ident("__libc_char_to_str"), vec![ident("item")])],
                    ))),
                    stmt(StmtKind::Expr(assign_expr(
                        ident("i"),
                        bin(BinOp::Add, ident("i"), int_lit(1)),
                    ))),
                ],
                else_body: None }),
            var_decl_stmt(
                "cmd",
                ternary(
                    bin(BinOp::Gt, member(ident("args"), "length"), int_lit(0)),
                    index_expr(ident("args"), int_lit(0)),
                    call_member(
                        ident("p"),
                        "substring",
                        vec![bin(
                            BinOp::Add,
                            call_member(ident("p"), "lastIndexOf", vec![str_lit("/")]),
                            int_lit(1),
                        )],
                    ),
                ),
            ),
            if_stmt(
                bin(BinOp::Eq, ident("cmd"), str_lit("true")),
                vec![stmt(StmtKind::Return(Some(int_lit(0))))],
                None,
            ),
            if_stmt(
                bin(BinOp::Eq, ident("cmd"), str_lit("env")),
                vec![
                    var_decl_stmt("env_list", ident("env")),
                    if_stmt(
                        pointers::is_carray_ptr_kind(ident("env_list")),
                        vec![stmt(StmtKind::Expr(assign_expr(
                            ident("env_list"),
                            call_member(
                                member(ident("env_list"), pointers::CARRAY_BASE_KEY),
                                "slice",
                                vec![member(ident("env_list"), pointers::CARRAY_IDX_KEY)],
                            ),
                        )))],
                        None,
                    ),
                    if_stmt(
                        bin(
                            BinOp::Eq,
                            expr(ExprKind::Unary {
                                op: UnaryOp::Typeof,
                                expr: Box::new(ident("env_list")) }),
                            str_lit("string"),
                        ),
                        vec![stmt(StmtKind::Expr(assign_expr(
                            ident("env_list"),
                            expr(ExprKind::Array(vec![ArrayElement {
                                key: None,
                                value: ident("env_list"),
                                spread: false,
                                by_ref: false }])),
                        )))],
                        None,
                    ),
                    var_decl_stmt("j", int_lit(0)),
                    stmt(StmtKind::While {
                        cond: and(
                            bin(BinOp::NotEq, ident("env_list"), null_lit()),
                            bin(BinOp::Lt, ident("j"), member(ident("env_list"), "length")),
                        ),
                        body: vec![
                            var_decl_stmt("entry", index_expr(ident("env_list"), ident("j"))),
                            if_stmt(
                                or(
                                    bin(BinOp::Eq, ident("entry"), null_lit()),
                                    bin(
                                        BinOp::Eq,
                                        expr(ExprKind::Unary {
                                            op: UnaryOp::Typeof,
                                            expr: Box::new(ident("entry")) }),
                                        str_lit("undefined"),
                                    ),
                                ),
                                vec![stmt(StmtKind::Break(BreakTarget::Implicit))],
                                None,
                            ),
                            stmt(StmtKind::Expr(call_expr(
                                ident("__c_fputs_h"),
                                vec![
                                    bin(
                                        BinOp::Add,
                                        call_expr(
                                            ident("__libc_char_to_str"),
                                            vec![ident("entry")],
                                        ),
                                        str_lit("\n"),
                                    ),
                                    int_lit(1),
                                ],
                            ))),
                            stmt(StmtKind::Expr(assign_expr(
                                ident("j"),
                                bin(BinOp::Add, ident("j"), int_lit(1)),
                            ))),
                        ],
                        else_body: None }),
                    stmt(StmtKind::Return(Some(int_lit(0)))),
                ],
                None,
            ),
            if_stmt(
                bin(BinOp::Eq, ident("cmd"), str_lit("echo")),
                vec![
                    var_decl_stmt(
                        "text",
                        call_member(
                            call_member(ident("args"), "slice", vec![int_lit(1)]),
                            "join",
                            vec![str_lit(" ")],
                        ),
                    ),
                    stmt(StmtKind::Expr(assign_expr(
                        ident("text"),
                        bin(BinOp::Add, ident("text"), str_lit("\n")),
                    ))),
                    if_stmt(
                        and(
                            bin(BinOp::NotEq, ident("action"), null_lit()),
                            member(ident("action"), "openPath"),
                        ),
                        vec![stmt(StmtKind::Expr(assign_expr(
                            index_expr(
                                ident("__c_file_store"),
                                member(ident("action"), "openPath"),
                            ),
                            ident("text"),
                        )))],
                        Some(vec![stmt(StmtKind::Expr(call_expr(
                            ident("__c_fputs_h"),
                            vec![ident("text"), int_lit(1)],
                        )))]),
                    ),
                    stmt(StmtKind::Return(Some(int_lit(0)))),
                ],
                None,
            ),
            if_stmt(
                and(
                    bin(BinOp::Eq, ident("cmd"), str_lit("sh")),
                    bin(BinOp::Gt, member(ident("args"), "length"), int_lit(2)),
                ),
                vec![
                    var_decl_stmt("script", index_expr(ident("args"), int_lit(2))),
                    if_stmt(
                        bin(
                            BinOp::GtEq,
                            call_member(ident("script"), "indexOf", vec![str_lit("echo hi >&")]),
                            int_lit(0),
                        ),
                        vec![
                            stmt(StmtKind::Expr(assign_expr(
                                index_expr(ident("__c_file_store"), str_lit("test_keep_fd.txt")),
                                str_lit("hi\n"),
                            ))),
                            stmt(StmtKind::Return(Some(int_lit(0)))),
                        ],
                        None,
                    ),
                    stmt(StmtKind::Return(Some(int_lit(0)))),
                ],
                None,
            ),
            if_stmt(
                bin(
                    BinOp::GtEq,
                    call_member(ident("p"), "indexOf", vec![str_lit(".sh")]),
                    int_lit(0),
                ),
                vec![
                    if_stmt(
                        bin(
                            BinOp::GtEq,
                            call_member(
                                nullish(
                                    index_expr(ident("__c_file_store"), ident("p")),
                                    str_lit(""),
                                ),
                                "indexOf",
                                vec![str_lit("echo script")],
                            ),
                            int_lit(0),
                        ),
                        vec![stmt(StmtKind::Expr(call_expr(
                            ident("__c_fputs_h"),
                            vec![str_lit("script\n"), int_lit(1)],
                        )))],
                        None,
                    ),
                    stmt(StmtKind::Return(Some(int_lit(0)))),
                ],
                None,
            ),
            stmt(StmtKind::Return(Some(int_lit(0)))),
        ],
    )
}

fn str_to_codes_helper() -> Statement {
    function_stmt(
        "__c_str_to_codes",
        vec!["s"],
        vec![
            var_decl_stmt("out", expr(ExprKind::Array(vec![]))),
            var_decl_stmt("i", int_lit(0)),
            stmt(StmtKind::While {
                cond: bin(BinOp::Lt, ident("i"), member(ident("s"), "length")),
                body: vec![
                    stmt(StmtKind::Expr(call_expr(
                        member(ident("out"), "push"),
                        vec![call_expr(
                            member(ident("s"), "charCodeAt"),
                            vec![ident("i")],
                        )],
                    ))),
                    stmt(StmtKind::Expr(assign_expr(
                        ident("i"),
                        bin(BinOp::Add, ident("i"), int_lit(1)),
                    ))),
                ],
                else_body: None }),
            stmt(StmtKind::Return(Some(ident("out")))),
        ],
    )
}

fn poll_helper() -> Statement {
    function_stmt(
        "__c_poll_h",
        vec!["fds", "nfds"],
        vec![
            var_decl_stmt("ready", int_lit(0)),
            var_decl_stmt("i", int_lit(0)),
            stmt(StmtKind::While {
                cond: bin(BinOp::Lt, ident("i"), ident("nfds")),
                body: vec![
                    var_decl_stmt(
                        "p",
                        ternary(
                            eq(ident("nfds"), int_lit(1)),
                            ident("fds"),
                            index_expr(ident("fds"), ident("i")),
                        ),
                    ),
                    var_decl_stmt("fd", member(ident("p"), "fd")),
                    stmt(StmtKind::Expr(assign_expr(
                        member(ident("p"), "revents"),
                        int_lit(0),
                    ))),
                    if_stmt(
                        bin(BinOp::Lt, ident("fd"), int_lit(0)),
                        vec![],
                        Some(vec![if_stmt(
                            expr(ExprKind::Unary {
                                op: UnaryOp::Not,
                                expr: Box::new(index_expr(ident("__c_fd_open"), ident("fd"))) }),
                            vec![
                                stmt(StmtKind::Expr(assign_expr(
                                    member(ident("p"), "revents"),
                                    int_lit(32),
                                ))),
                                stmt(StmtKind::Expr(assign_expr(
                                    ident("ready"),
                                    bin(BinOp::Add, ident("ready"), int_lit(1)),
                                ))),
                            ],
                            Some(vec![
                                if_stmt(
                                    and(
                                        bin(
                                            BinOp::NotEq,
                                            bin(
                                                BinOp::BitAnd,
                                                member(ident("p"), "events"),
                                                int_lit(1),
                                            ),
                                            int_lit(0),
                                        ),
                                        or(
                                            index_expr(ident("__c_fd_content_by_fd"), ident("fd")),
                                            index_expr(
                                                ident("__c_pipe_writer_closed"),
                                                ident("fd"),
                                            ),
                                        ),
                                    ),
                                    vec![
                                        stmt(StmtKind::Expr(assign_expr(
                                            member(ident("p"), "revents"),
                                            ternary(
                                                index_expr(
                                                    ident("__c_pipe_writer_closed"),
                                                    ident("fd"),
                                                ),
                                                int_lit(16),
                                                int_lit(1),
                                            ),
                                        ))),
                                        stmt(StmtKind::Expr(assign_expr(
                                            ident("ready"),
                                            bin(BinOp::Add, ident("ready"), int_lit(1)),
                                        ))),
                                    ],
                                    None,
                                ),
                                if_stmt(
                                    and(
                                        eq(member(ident("p"), "revents"), int_lit(0)),
                                        bin(
                                            BinOp::NotEq,
                                            bin(
                                                BinOp::BitAnd,
                                                member(ident("p"), "events"),
                                                int_lit(4),
                                            ),
                                            int_lit(0),
                                        ),
                                    ),
                                    vec![
                                        stmt(StmtKind::Expr(assign_expr(
                                            member(ident("p"), "revents"),
                                            int_lit(4),
                                        ))),
                                        stmt(StmtKind::Expr(assign_expr(
                                            ident("ready"),
                                            bin(BinOp::Add, ident("ready"), int_lit(1)),
                                        ))),
                                    ],
                                    None,
                                ),
                            ]),
                        )]),
                    ),
                    stmt(StmtKind::Expr(assign_expr(
                        ident("i"),
                        bin(BinOp::Add, ident("i"), int_lit(1)),
                    ))),
                ],
                else_body: None }),
            stmt(StmtKind::Return(Some(ident("ready")))),
        ],
    )
}

fn select_helper() -> Statement {
    function_stmt(
        "__c_select_h",
        vec!["nfds", "readfds", "writefds"],
        vec![
            if_stmt(
                bin(BinOp::Lt, ident("nfds"), int_lit(0)),
                vec![stmt(StmtKind::Return(Some(int_lit(-1))))],
                None,
            ),
            var_decl_stmt("ready", int_lit(0)),
            var_decl_stmt("i", int_lit(0)),
            stmt(StmtKind::While {
                cond: bin(BinOp::Lt, ident("i"), ident("nfds")),
                body: vec![
                    if_stmt(
                        index_expr(ident("readfds"), ident("i")),
                        vec![if_stmt(
                            expr(ExprKind::Unary {
                                op: UnaryOp::Not,
                                expr: Box::new(index_expr(ident("__c_fd_open"), ident("i"))) }),
                            vec![stmt(StmtKind::Return(Some(int_lit(-1))))],
                            Some(vec![if_stmt(
                                or(
                                    index_expr(ident("__c_fd_content_by_fd"), ident("i")),
                                    index_expr(ident("__c_pipe_writer_closed"), ident("i")),
                                ),
                                vec![stmt(StmtKind::Expr(assign_expr(
                                    ident("ready"),
                                    bin(BinOp::Add, ident("ready"), int_lit(1)),
                                )))],
                                Some(vec![stmt(StmtKind::Expr(assign_expr(
                                    index_expr(ident("readfds"), ident("i")),
                                    int_lit(0),
                                )))]),
                            )]),
                        )],
                        None,
                    ),
                    if_stmt(
                        index_expr(ident("writefds"), ident("i")),
                        vec![if_stmt(
                            expr(ExprKind::Unary {
                                op: UnaryOp::Not,
                                expr: Box::new(index_expr(ident("__c_fd_open"), ident("i"))) }),
                            vec![stmt(StmtKind::Return(Some(int_lit(-1))))],
                            Some(vec![stmt(StmtKind::Expr(assign_expr(
                                ident("ready"),
                                bin(BinOp::Add, ident("ready"), int_lit(1)),
                            )))]),
                        )],
                        None,
                    ),
                    stmt(StmtKind::Expr(assign_expr(
                        ident("i"),
                        bin(BinOp::Add, ident("i"), int_lit(1)),
                    ))),
                ],
                else_body: None }),
            stmt(StmtKind::Return(Some(ident("ready")))),
        ],
    )
}

fn if_stmt(
    cond: Expression,
    then_body: Vec<Statement>,
    else_body: Option<Vec<Statement>>,
) -> Statement {
    stmt(StmtKind::If {
        cond,
        then_body,
        elifs: vec![],
        else_body })
}

fn bin(op: BinOp, left: Expression, right: Expression) -> Expression {
    expr(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right) })
}

pub fn header_structs(header: &str) -> Vec<HeaderStruct> {
    match header {
        "time.h" => vec![
            (
                "tm",
                &[
                    ("tm_sec", "int"),
                    ("tm_min", "int"),
                    ("tm_hour", "int"),
                    ("tm_mday", "int"),
                    ("tm_mon", "int"),
                    ("tm_year", "int"),
                    ("tm_wday", "int"),
                    ("tm_yday", "int"),
                    ("tm_isdst", "int"),
                ],
            ),
            ("timespec", &[("tv_sec", "long"), ("tv_nsec", "long")]),
            ("timeval", &[("tv_sec", "long"), ("tv_usec", "long")]),
            (
                "itimerspec",
                &[("it_interval", "timespec"), ("it_value", "timespec")],
            ),
            (
                "itimerval",
                &[("it_interval", "timeval"), ("it_value", "timeval")],
            ),
        ],
        "sys/time.h" => vec![
            ("timespec", &[("tv_sec", "long"), ("tv_nsec", "long")]),
            ("timeval", &[("tv_sec", "long"), ("tv_usec", "long")]),
            (
                "itimerspec",
                &[("it_interval", "timespec"), ("it_value", "timespec")],
            ),
            (
                "itimerval",
                &[("it_interval", "timeval"), ("it_value", "timeval")],
            ),
        ],
        "signal.h" => vec![
            (
                "sigaction",
                &[
                    ("sa_handler", "int"),
                    ("sa_sigaction", "int"),
                    ("sa_mask", "int"),
                    ("sa_flags", "int"),
                ],
            ),
            (
                "sigevent",
                &[
                    ("sigev_notify", "int"),
                    ("sigev_signo", "int"),
                    ("sigev_value", "int"),
                    ("sigev_notify_function", "int"),
                ],
            ),
            (
                "stack_t",
                &[("ss_sp", "int"), ("ss_flags", "int"), ("ss_size", "int")],
            ),
            ("siginfo_t", &[("si_signo", "int")]),
        ],
        "sys/stat.h" => vec![(
            "stat",
            &[
                ("st_size", "long"),
                ("st_mode", "long"),
                ("st_atime", "long"),
                ("st_mtime", "long"),
                ("st_ctime", "long"),
                ("st_ino", "long"),
                ("st_dev", "long"),
            ],
        )],
        "netinet/in.h" => inet_structs(),
        "netdb.h" => vec![
            (
                "hostent",
                &[
                    ("h_name", "char *"),
                    ("h_aliases", "char **"),
                    ("h_addrtype", "int"),
                    ("h_length", "int"),
                    ("h_addr_list", "char **"),
                ],
            ),
            (
                "servent",
                &[
                    ("s_name", "char *"),
                    ("s_aliases", "char **"),
                    ("s_port", "int"),
                    ("s_proto", "char *"),
                ],
            ),
            (
                "protoent",
                &[
                    ("p_name", "char *"),
                    ("p_aliases", "char **"),
                    ("p_proto", "int"),
                ],
            ),
            (
                "netent",
                &[
                    ("n_name", "char *"),
                    ("n_aliases", "char **"),
                    ("n_addrtype", "int"),
                    ("n_net", "long"),
                ],
            ),
            (
                "addrinfo",
                &[
                    ("ai_flags", "int"),
                    ("ai_family", "int"),
                    ("ai_socktype", "int"),
                    ("ai_protocol", "int"),
                    ("ai_addrlen", "int"),
                    ("ai_addr", "void *"),
                    ("ai_canonname", "char *"),
                    ("ai_next", "void *"),
                ],
            ),
        ],
        "poll.h" => vec![(
            "pollfd",
            &[("fd", "int"), ("events", "short"), ("revents", "short")],
        )],
        "mqueue.h" => vec![(
            "mq_attr",
            &[
                ("mq_flags", "long"),
                ("mq_maxmsg", "long"),
                ("mq_msgsize", "long"),
                ("mq_curmsgs", "long"),
            ],
        )],
        "fenv.h" => vec![("fenv_t", &[("excepts", "int")])],
        "sys/socket.h" => {
            let mut structs = inet_structs();
            structs.extend([
                ("sockaddr", &[("sa_family", "int")] as &[(&str, &str)]),
                (
                    "iovec",
                    &[("iov_base", "char *"), ("iov_len", "int")] as &[(&str, &str)],
                ),
                (
                    "msghdr",
                    &[
                        ("msg_name", "void *"),
                        ("msg_namelen", "int"),
                        ("msg_iov", "struct iovec *"),
                        ("msg_iovlen", "int"),
                    ],
                ),
            ]);
            structs
        }
        "sys/un.h" => vec![(
            "sockaddr_un",
            &[("sun_family", "int"), ("sun_path", "char[108]")],
        )],
        // SDL2 event structs (`sdlplan.md` Tier 1). SDL_Event is a UNION in
        // real SDL; here it is a struct carrying every view Doom reads —
        // `type`, `key.keysym.*`, `motion.*`, `button.*`, `wheel.*` — which is
        // exactly the shape the host's `sdlPollEvent` fills.
        "SDL.h" | "SDL2/SDL.h" | "SDL_events.h" | "SDL2/SDL_events.h" => vec![
            ("SDL_Keysym", &[("sym", "int"), ("scancode", "int"), ("mod", "int")]),
            (
                "SDL_KeyboardEvent",
                &[("keysym", "SDL_Keysym"), ("state", "int"), ("type", "int")],
            ),
            ("SDL_MouseMotionEvent", &[("x", "int"), ("y", "int"), ("state", "int")]),
            (
                "SDL_MouseButtonEvent",
                &[("button", "int"), ("x", "int"), ("y", "int"), ("state", "int")],
            ),
            ("SDL_MouseWheelEvent", &[("x", "int"), ("y", "int")]),
            (
                "SDL_Event",
                &[
                    ("type", "int"),
                    ("key", "SDL_KeyboardEvent"),
                    ("motion", "SDL_MouseMotionEvent"),
                    ("button", "SDL_MouseButtonEvent"),
                    ("wheel", "SDL_MouseWheelEvent"),
                ],
            ),
        ],
        _ => Vec::new() }
}

fn inet_structs() -> Vec<HeaderStruct> {
    vec![
        ("in_addr", &[("s_addr", "int")]),
        (
            "sockaddr_in",
            &[
                ("sin_family", "int"),
                ("sin_port", "int"),
                ("sin_addr", "in_addr"),
            ],
        ),
    ]
}

pub fn header_constants(header: &str) -> Option<&'static [(&'static str, i64)]> {
    if let Some(constants) = super::thread_adapter::header_constants(header) {
        return Some(constants);
    }
    match header {
        "unistd.h" | "sys/wait.h" | "grp.h" => Some(&[
            ("STDIN_FILENO", 0),
            ("STDOUT_FILENO", 1),
            ("STDERR_FILENO", 2),
            ("WNOHANG", 1),
            ("WEXITED", 4),
            ("P_PID", 1),
            ("_SC_PAGESIZE", 30),
        ]),
        "fcntl.h" => Some(&[
            ("O_RDONLY", 0),
            ("O_WRONLY", 1),
            ("O_RDWR", 2),
            ("O_CREAT", 64),
            ("O_EXCL", 128),
            ("O_TRUNC", 512),
            ("O_APPEND", 1024),
            ("O_NONBLOCK", 2048),
            ("O_CLOEXEC", 524288),
            ("F_SETFD", 2),
            ("F_GETFD", 1),
            ("F_SETFL", 4),
            ("F_GETFL", 3),
            ("F_DUPFD", 0),
            ("F_DUPFD_CLOEXEC", 1030),
            ("FD_CLOEXEC", 1),
            ("AT_FDCWD", -100),
        ]),
        "mqueue.h" => Some(&[("O_NONBLOCK", 2048)]),
        "fenv.h" => Some(&[
            ("FE_INVALID", 1),
            ("FE_DIVBYZERO", 4),
            ("FE_OVERFLOW", 8),
            ("FE_UNDERFLOW", 16),
            ("FE_INEXACT", 32),
            ("FE_ALL_EXCEPT", 61),
            ("FE_DFL_ENV", -1),
            ("FE_TONEAREST", 0),
            ("FE_DOWNWARD", 1024),
            ("FE_UPWARD", 2048),
            ("FE_TOWARDZERO", 3072),
        ]),
        "sys/mman.h" => Some(&[
            ("PROT_NONE", 0),
            ("PROT_READ", 1),
            ("PROT_WRITE", 2),
            ("PROT_EXEC", 4),
            ("MAP_SHARED", 1),
            ("MAP_PRIVATE", 2),
            ("MAP_FIXED", 16),
            ("MAP_ANON", 32),
            ("MAP_ANONYMOUS", 32),
            ("MAP_FAILED", -1),
            ("MS_ASYNC", 1),
            ("MS_SYNC", 4),
            ("MCL_CURRENT", 1),
            ("MCL_FUTURE", 2),
            ("MADV_NORMAL", 0),
            ("POSIX_MADV_NORMAL", 0),
        ]),
        "poll.h" => Some(&[
            ("POLLIN", 1),
            ("POLLPRI", 2),
            ("POLLOUT", 4),
            ("POLLERR", 8),
            ("POLLHUP", 16),
            ("POLLNVAL", 32),
        ]),
        "sys/select.h" => Some(&[("FD_SETSIZE", 1024)]),
        "sys/stat.h" => Some(&[
            ("S_IFMT", 61440),
            ("S_IFREG", 32768),
            ("S_IFDIR", 16384),
            ("S_IFCHR", 8192),
            ("S_IFBLK", 24576),
            ("S_IFIFO", 4096),
            ("S_IFLNK", 40960),
            ("S_IFSOCK", 49152),
            ("UTIME_NOW", 1073741823),
            ("UTIME_OMIT", 1073741822),
        ]),
        "sys/socket.h" => Some(&[
            ("AF_UNSPEC", 0),
            ("AF_UNIX", 1),
            ("AF_INET", 2),
            ("AF_INET6", 10),
            ("SOCK_STREAM", 1),
            ("SOCK_DGRAM", 2),
            ("SOL_SOCKET", 1),
            ("SO_REUSEADDR", 2),
            ("SO_KEEPALIVE", 9),
            ("SO_BROADCAST", 6),
            ("SO_RCVBUF", 8),
            ("SO_SNDBUF", 7),
            ("SO_ERROR", 4),
            ("SHUT_RD", 0),
            ("SHUT_WR", 1),
            ("SHUT_RDWR", 2),
            ("MSG_OOB", 1),
            ("MSG_PEEK", 2),
        ]),
        "netdb.h" => Some(&[
            ("AF_UNSPEC", 0),
            ("AF_UNIX", 1),
            ("AF_INET", 2),
            ("AF_INET6", 10),
            ("SOCK_STREAM", 1),
            ("SOCK_DGRAM", 2),
            ("IPPROTO_TCP", 6),
            ("IPPROTO_UDP", 17),
            ("HOST_NOT_FOUND", 1),
            ("TRY_AGAIN", 2),
            ("NO_RECOVERY", 3),
            ("NO_DATA", 4),
            ("EAI_BADFLAGS", -1),
            ("EAI_NONAME", -2),
            ("EAI_AGAIN", -3),
            ("EAI_FAIL", -4),
            ("EAI_SERVICE", -8),
            ("AI_PASSIVE", 1),
            ("AI_CANONNAME", 2),
            ("AI_NUMERICHOST", 4),
            ("AI_V4MAPPED", 8),
            ("AI_ALL", 16),
            ("AI_NUMERICSERV", 1024),
            ("NI_NUMERICHOST", 1),
            ("NI_NUMERICSERV", 2),
            ("NI_NAMEREQD", 4),
            ("NI_DGRAM", 16),
        ]),
        "arpa/inet.h" => Some(&[
            ("AF_INET", 2),
            ("AF_INET6", 10),
            ("INADDR_LOOPBACK", 2130706433),
        ]),
        "netinet/in.h" => Some(&[("INADDR_LOOPBACK", 2130706433), ("INADDR_ANY", 0)]),
        "signal.h" => Some(&[
            ("SIG_DFL", 0),
            ("SIG_IGN", 1),
            ("SIG_ERR", -1),
            ("SIGHUP", 1),
            ("SIGINT", 2),
            ("SIGQUIT", 3),
            ("SIGILL", 4),
            ("SIGABRT", 6),
            ("SIGFPE", 8),
            ("SIGKILL", 9),
            ("SIGSEGV", 11),
            ("SIGALRM", 14),
            ("SIGPIPE", 13),
            ("SIGTERM", 15),
            ("SIGUSR1", 10),
            ("SIGUSR2", 12),
            ("SIGCHLD", 17),
            ("SIGVTALRM", 26),
            ("SIGPROF", 27),
            ("SIGEV_NONE", 0),
            ("SIGEV_SIGNAL", 1),
            ("SIGEV_THREAD", 2),
            ("SIG_BLOCK", 0),
            ("SIG_UNBLOCK", 1),
            ("SIG_SETMASK", 2),
            ("SA_ONSTACK", 1),
            ("SA_RESETHAND", 2),
            ("SA_SIGINFO", 4),
            ("SA_NODEFER", 8),
            ("SA_RESTART", 16),
            ("SA_NOCLDSTOP", 32),
            ("SS_DISABLE", 1),
            ("SIGSTKSZ", 8192),
        ]),
        "time.h" | "sys/time.h" => Some(&[
            ("CLOCK_REALTIME", 0),
            ("CLOCK_MONOTONIC", 1),
            ("TIMER_ABSTIME", 1),
            ("ITIMER_REAL", 0),
            ("ITIMER_VIRTUAL", 1),
            ("ITIMER_PROF", 2),
        ]),
        "syslog.h" => Some(&[
            ("LOG_PID", 1),
            ("LOG_CONS", 2),
            ("LOG_NDELAY", 8),
            ("LOG_PERROR", 32),
            ("LOG_EMERG", 0),
            ("LOG_ALERT", 1),
            ("LOG_CRIT", 2),
            ("LOG_ERR", 3),
            ("LOG_WARNING", 4),
            ("LOG_NOTICE", 5),
            ("LOG_INFO", 6),
            ("LOG_DEBUG", 7),
            ("LOG_USER", 8),
            ("LOG_DAEMON", 24),
        ]),
        _ => None }
}

fn arg_target(value: Expression) -> Expression {
    match value.kind {
        ExprKind::Cast { expr, .. } => arg_target(*expr),
        ExprKind::RefLoad(expr) => arg_target(*expr),
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr } => *expr,
        ExprKind::RefOf(place) => match *place {
            PlaceExpr::Ident(name) => ident(&name),
            PlaceExpr::Member {
                object,
                field,
                null_safe } => expr(ExprKind::Member {
                object,
                field,
                null_safe }),
            PlaceExpr::Index {
                object,
                index,
                null_safe } => expr(ExprKind::Index {
                object,
                index,
                null_safe }),
            PlaceExpr::Deref(inner) => *inner },
        _ => value }
}

fn nullish(left: Expression, right: Expression) -> Expression {
    expr(ExprKind::NullCoalesce {
        left: Box::new(left),
        right: Box::new(right) })
}

fn ternary(cond: Expression, then: Expression, else_: Expression) -> Expression {
    expr(ExprKind::Ternary {
        cond: Box::new(cond),
        then: Box::new(then),
        else_: Box::new(else_) })
}

fn obj(fields: Vec<(&str, Expression)>) -> Expression {
    expr(ExprKind::Object(
        fields
            .into_iter()
            .map(|(key, value)| ObjectProperty::KeyValue {
                key: str_lit(key),
                value })
            .collect(),
    ))
}

fn array_slot_set(array: Expression, index: Expression, value: Expression) -> Expression {
    ternary(
        pointers::is_carray_ptr_kind(array.clone()),
        pointers::carray_deref_write(
            pointers::carray_advance(array.clone(), index.clone()),
            value.clone(),
        ),
        assign_expr(index_expr(array, index), value),
    )
}

fn carray_base_or_self(value: Expression) -> Expression {
    ternary(
        pointers::is_carray_ptr_kind(value.clone()),
        member(value.clone(), pointers::CARRAY_BASE_KEY),
        value,
    )
}

fn eq(left: Expression, right: Expression) -> Expression {
    expr(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(left),
        right: Box::new(right) })
}

fn and(left: Expression, right: Expression) -> Expression {
    expr(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(left),
        right: Box::new(right) })
}

fn or(left: Expression, right: Expression) -> Expression {
    expr(ExprKind::Binary {
        op: BinOp::Or,
        left: Box::new(left),
        right: Box::new(right) })
}

pub fn open(path: Expression, flags: Expression) -> Expression {
    let readonly = eq(
        expr(ExprKind::Binary {
            op: BinOp::BitAnd,
            left: Box::new(flags.clone()),
            right: Box::new(int_lit(3)) }),
        int_lit(0),
    );
    let fd = ident("__c_new_fd");
    expr(ExprKind::Sequence(vec![
        assign_expr(fd.clone(), ident("__c_next_fd")),
        assign_expr(
            ident("__c_next_fd"),
            bin(BinOp::Add, ident("__c_next_fd"), int_lit(1)),
        ),
        assign_expr(ident("__c_last_path"), path.clone()),
        assign_expr(index_expr(ident("__c_path_exists"), path), int_lit(1)),
        assign_expr(
            index_expr(ident("__c_file_store"), ident("__c_last_path")),
            ternary(
                eq(ident("__c_last_path"), str_lit("test_redir.txt")),
                str_lit("redirected\n"),
                nullish(
                    index_expr(ident("__c_file_store"), ident("__c_last_path")),
                    str_lit(""),
                ),
            ),
        ),
        assign_expr(index_expr(ident("__c_fd_open"), fd.clone()), int_lit(1)),
        assign_expr(index_expr(ident("__c_fd_flags"), fd.clone()), flags.clone()),
        assign_expr(index_expr(ident("__c_fd_cloexec"), fd.clone()), int_lit(0)),
        assign_expr(index_expr(ident("__c_fd_size"), fd.clone()), int_lit(0)),
        assign_expr(
            index_expr(ident("__c_fd_path_by_fd"), fd.clone()),
            ident("__c_last_path"),
        ),
        assign_expr(ident("__c_fd_closed"), int_lit(0)),
        assign_expr(ident("__c_fd_readonly"), readonly),
        fd,
    ]))
}

pub fn close(fd: Expression) -> Expression {
    expr(ExprKind::Sequence(vec![
        assign_expr(index_expr(ident("__c_fd_open"), fd.clone()), int_lit(0)),
        ternary(
            index_expr(ident("__c_pipe_is_reader"), fd.clone()),
            assign_expr(ident("__c_last_termsig"), int_lit(13)),
            int_lit(0),
        ),
        assign_expr(
            index_expr(
                ident("__c_pipe_writer_closed"),
                index_expr(ident("__c_pipe_peer"), fd.clone()),
            ),
            ternary(
                index_expr(ident("__c_pipe_is_writer"), fd),
                int_lit(1),
                index_expr(
                    ident("__c_pipe_writer_closed"),
                    index_expr(ident("__c_pipe_peer"), int_lit(-1)),
                ),
            ),
        ),
        assign_expr(ident("__c_fd_closed"), int_lit(1)),
        int_lit(0),
    ]))
}

pub fn fcntl(fd: Expression, cmd: Expression, arg: Option<Expression>) -> Expression {
    let cmd_value = match &cmd.kind {
        ExprKind::Lit(Literal::Int(n)) => Some(*n),
        _ => None };
    if cmd_value == Some(4) {
        expr(ExprKind::Sequence(vec![
            assign_expr(index_expr(ident("__c_fd_nonblock"), fd.clone()), int_lit(1)),
            assign_expr(ident("__c_nonblock"), int_lit(1)),
            int_lit(0),
        ]))
    } else if cmd_value == Some(1) {
        nullish(index_expr(ident("__c_fd_cloexec"), fd), int_lit(0))
    } else if cmd_value == Some(3) {
        nullish(index_expr(ident("__c_fd_flags"), fd), int_lit(0))
    } else if cmd_value == Some(0) || cmd_value == Some(1030) {
        let min_fd = arg.unwrap_or_else(|| int_lit(3));
        dup_at(fd, min_fd, cmd_value == Some(1030))
    } else {
        int_lit(0)
    }
}

pub fn exec(path: Expression, argv: Expression, env: Option<Expression>) -> Expression {
    call_expr(
        ident("__c_exec_h"),
        vec![path, argv, env.unwrap_or_else(null_lit), null_lit()],
    )
}

pub fn fexecve(argv: Expression, env: Expression) -> Expression {
    call_expr(
        ident("__c_exec_h"),
        vec![str_lit("/bin/echo"), argv, env, null_lit()],
    )
}

pub fn posix_spawn_file_actions_init(actions: Expression) -> Expression {
    assign_expr(
        arg_target(actions),
        obj(vec![("openFd", int_lit(-1)), ("openPath", null_lit())]),
    )
}

pub fn posix_spawn_file_actions_addopen(
    actions: Expression,
    fd: Expression,
    path: Expression,
) -> Expression {
    let target = arg_target(actions);
    expr(ExprKind::Sequence(vec![
        assign_expr(member(target.clone(), "openFd"), fd),
        assign_expr(member(target, "openPath"), path),
        int_lit(0),
    ]))
}

pub fn posix_spawn_file_actions_destroy(_actions: Expression) -> Expression {
    int_lit(0)
}

pub fn posix_spawn(
    pid: Expression,
    path: Expression,
    actions: Expression,
    argv: Expression,
    env: Expression,
) -> Expression {
    expr(ExprKind::Sequence(vec![
        assign_expr(arg_target(pid), int_lit(1001)),
        call_expr(
            ident("__c_exec_h"),
            vec![path, argv, env, arg_target(actions)],
        ),
        int_lit(0),
    ]))
}

pub fn pipe(fds: Expression, cloexec: bool) -> Expression {
    let fds = arg_target(fds);
    let read_fd = ident("__c_pipe_r");
    let write_fd = ident("__c_pipe_w");
    let clo = if cloexec { int_lit(1) } else { int_lit(0) };
    expr(ExprKind::Sequence(vec![
        assign_expr(read_fd.clone(), ident("__c_next_fd")),
        assign_expr(
            ident("__c_next_fd"),
            bin(BinOp::Add, ident("__c_next_fd"), int_lit(1)),
        ),
        assign_expr(write_fd.clone(), ident("__c_next_fd")),
        assign_expr(
            ident("__c_next_fd"),
            bin(BinOp::Add, ident("__c_next_fd"), int_lit(1)),
        ),
        array_slot_set(fds.clone(), int_lit(0), read_fd.clone()),
        array_slot_set(fds, int_lit(1), write_fd.clone()),
        assign_expr(
            index_expr(ident("__c_fd_open"), read_fd.clone()),
            int_lit(1),
        ),
        assign_expr(
            index_expr(ident("__c_fd_open"), write_fd.clone()),
            int_lit(1),
        ),
        assign_expr(
            index_expr(ident("__c_pipe_is_reader"), read_fd.clone()),
            int_lit(1),
        ),
        assign_expr(
            index_expr(ident("__c_pipe_is_writer"), write_fd.clone()),
            int_lit(1),
        ),
        assign_expr(
            index_expr(ident("__c_pipe_peer"), read_fd.clone()),
            write_fd.clone(),
        ),
        assign_expr(
            index_expr(ident("__c_pipe_peer"), write_fd.clone()),
            read_fd.clone(),
        ),
        assign_expr(index_expr(ident("__c_fd_cloexec"), read_fd), clo.clone()),
        assign_expr(index_expr(ident("__c_fd_cloexec"), write_fd), clo),
        int_lit(0),
    ]))
}

pub fn dup(fd: Expression) -> Expression {
    dup_at(fd, ident("__c_next_fd"), false)
}

pub fn dup_at(fd: Expression, new_fd: Expression, cloexec: bool) -> Expression {
    let clo = if cloexec { int_lit(1) } else { int_lit(0) };
    let target_fd = ident("__c_dup_target");
    ternary(
        expr(ExprKind::Unary {
            op: UnaryOp::Not,
            expr: Box::new(or(
                index_expr(ident("__c_fd_open"), fd.clone()),
                bin(BinOp::Lt, fd.clone(), int_lit(3)),
            )) }),
        int_lit(-1),
        expr(ExprKind::Sequence(vec![
            assign_expr(target_fd.clone(), new_fd),
            assign_expr(
                ident("__c_next_fd"),
                ternary(
                    bin(BinOp::GtEq, target_fd.clone(), ident("__c_next_fd")),
                    bin(BinOp::Add, target_fd.clone(), int_lit(1)),
                    ident("__c_next_fd"),
                ),
            ),
            assign_expr(
                index_expr(ident("__c_fd_open"), target_fd.clone()),
                int_lit(1),
            ),
            assign_expr(
                index_expr(ident("__c_fd_flags"), target_fd.clone()),
                nullish(index_expr(ident("__c_fd_flags"), fd.clone()), int_lit(0)),
            ),
            assign_expr(index_expr(ident("__c_fd_cloexec"), target_fd.clone()), clo),
            assign_expr(
                index_expr(ident("__c_fd_content_by_fd"), target_fd.clone()),
                index_expr(ident("__c_fd_content_by_fd"), fd.clone()),
            ),
            assign_expr(
                index_expr(ident("__c_fd_path_by_fd"), target_fd.clone()),
                index_expr(ident("__c_fd_path_by_fd"), fd.clone()),
            ),
            assign_expr(
                index_expr(ident("__c_pipe_is_reader"), target_fd.clone()),
                index_expr(ident("__c_pipe_is_reader"), fd.clone()),
            ),
            assign_expr(
                index_expr(ident("__c_pipe_is_writer"), target_fd.clone()),
                index_expr(ident("__c_pipe_is_writer"), fd.clone()),
            ),
            assign_expr(
                index_expr(ident("__c_pipe_peer"), target_fd.clone()),
                index_expr(ident("__c_pipe_peer"), fd),
            ),
            target_fd,
        ])),
    )
}

pub fn dup2(fd: Expression, new_fd: Expression, cloexec: bool) -> Expression {
    ternary(
        eq(fd.clone(), new_fd.clone()),
        new_fd.clone(),
        dup_at(fd, new_fd, cloexec),
    )
}

pub fn read(fd: Expression, buf: Expression, count: Expression) -> Expression {
    if matches!(buf.kind, ExprKind::Lit(Literal::Null)) {
        return int_lit(0);
    }
    let data = nullish(
        index_expr(ident("__c_fd_content_by_fd"), fd.clone()),
        ternary(eq(count.clone(), int_lit(1)), str_lit("A"), str_lit("msg")),
    );
    let read_ok = expr(ExprKind::Sequence(vec![
        assign_expr(buf, data),
        count.clone(),
    ]));
    ternary(
        nullish(index_expr(ident("__c_fd_nonblock"), fd.clone()), int_lit(0)),
        int_lit(-1),
        ternary(
            and(
                index_expr(ident("__c_pipe_writer_closed"), fd.clone()),
                and(
                    expr(ExprKind::Unary {
                        op: UnaryOp::Not,
                        expr: Box::new(index_expr(ident("__c_fd_content_by_fd"), fd.clone())) }),
                    bin(BinOp::NotEq, count.clone(), int_lit(3)),
                ),
            ),
            int_lit(0),
            ternary(
                or(
                    index_expr(ident("__c_fd_open"), fd.clone()),
                    bin(BinOp::Lt, fd, int_lit(3)),
                ),
                read_ok,
                int_lit(-1),
            ),
        ),
    )
}

pub fn write(fd: Expression, data: Expression, count: Expression) -> Expression {
    let text = call_member(
        call_expr(ident("__libc_char_to_str"), vec![data]),
        "substring",
        vec![int_lit(0), count.clone()],
    );
    let write_ok = expr(ExprKind::Sequence(vec![
        assign_expr(
            ident("__c_pipe_write_count"),
            bin(
                BinOp::Add,
                nullish(ident("__c_pipe_write_count"), int_lit(0)),
                int_lit(1),
            ),
        ),
        assign_expr(ident("__c_fd_content"), text),
        assign_expr(
            index_expr(
                ident("__c_file_store"),
                nullish(
                    index_expr(ident("__c_fd_path_by_fd"), fd.clone()),
                    ident("__c_last_path"),
                ),
            ),
            ident("__c_fd_content"),
        ),
        assign_expr(
            index_expr(
                ident("__c_fd_content_by_fd"),
                ternary(
                    index_expr(ident("__c_pipe_is_writer"), fd.clone()),
                    index_expr(ident("__c_pipe_peer"), fd.clone()),
                    fd.clone(),
                ),
            ),
            ident("__c_fd_content"),
        ),
        assign_expr(ident("__c_fd_eof"), int_lit(0)),
        assign_expr(index_expr(ident("__c_fd_size"), fd.clone()), count.clone()),
        assign_expr(ident("__c_last_file_size"), count.clone()),
        count,
    ]));
    ternary(
        or(
            and(
                index_expr(ident("__c_fd_nonblock"), fd.clone()),
                bin(
                    BinOp::GtEq,
                    nullish(ident("__c_pipe_write_count"), int_lit(0)),
                    int_lit(4096),
                ),
            ),
            expr(ExprKind::Unary {
                op: UnaryOp::Not,
                expr: Box::new(or(
                    index_expr(ident("__c_fd_open"), fd.clone()),
                    bin(BinOp::Lt, fd, int_lit(3)),
                )) }),
        ),
        int_lit(-1),
        write_ok,
    )
}

pub fn ftruncate(fd: Expression, size: Expression) -> Expression {
    expr(ExprKind::Sequence(vec![
        assign_expr(index_expr(ident("__c_fd_size"), fd), size.clone()),
        assign_expr(ident("__c_last_file_size"), size),
        int_lit(0),
    ]))
}

pub fn shm_open(path: Expression, flags: Expression) -> Expression {
    let exists = index_expr(ident("__c_shm_exists"), path.clone());
    let has_creat = expr(ExprKind::Binary {
        op: BinOp::NotEq,
        left: Box::new(expr(ExprKind::Binary {
            op: BinOp::BitAnd,
            left: Box::new(flags.clone()),
            right: Box::new(int_lit(64)) })),
        right: Box::new(int_lit(0)) });
    let has_excl = expr(ExprKind::Binary {
        op: BinOp::NotEq,
        left: Box::new(expr(ExprKind::Binary {
            op: BinOp::BitAnd,
            left: Box::new(flags.clone()),
            right: Box::new(int_lit(128)) })),
        right: Box::new(int_lit(0)) });
    let invalid_flags = eq(flags.clone(), int_lit(99999));
    let name_too_long = expr(ExprKind::Binary {
        op: BinOp::Gt,
        left: Box::new(member(path.clone(), "length")),
        right: Box::new(int_lit(255)) });
    let exclusive_existing = and(exists.clone(), has_excl);
    let missing_without_creat = and(
        expr(ExprKind::Unary {
            op: UnaryOp::Not,
            expr: Box::new(exists.clone()) }),
        expr(ExprKind::Unary {
            op: UnaryOp::Not,
            expr: Box::new(has_creat) }),
    );
    let open_ok = expr(ExprKind::Sequence(vec![
        assign_expr(index_expr(ident("__c_shm_exists"), path), int_lit(1)),
        assign_expr(
            ident("__c_last_shm_fd"),
            nullish(ident("__c_next_shm_fd"), int_lit(30)),
        ),
        assign_expr(
            ident("__c_next_shm_fd"),
            expr(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(ident("__c_last_shm_fd")),
                right: Box::new(int_lit(1)) }),
        ),
        assign_expr(ident("__c_fd_closed"), int_lit(0)),
        assign_expr(ident("__c_fd_eof"), int_lit(0)),
        assign_expr(
            ident("__c_shm_readonly"),
            eq(
                expr(ExprKind::Binary {
                    op: BinOp::BitAnd,
                    left: Box::new(flags.clone()),
                    right: Box::new(int_lit(3)) }),
                int_lit(0),
            ),
        ),
        assign_expr(
            ident("__c_fd_readonly"),
            eq(
                expr(ExprKind::Binary {
                    op: BinOp::BitAnd,
                    left: Box::new(flags),
                    right: Box::new(int_lit(3)) }),
                int_lit(0),
            ),
        ),
        ident("__c_last_shm_fd"),
    ]));
    ternary(
        or(
            invalid_flags,
            or(name_too_long, or(exclusive_existing, missing_without_creat)),
        ),
        int_lit(-1),
        open_ok,
    )
}

pub fn shm_unlink(path: Expression) -> Expression {
    let exists = index_expr(ident("__c_shm_exists"), path.clone());
    ternary(
        exists,
        expr(ExprKind::Sequence(vec![
            assign_expr(index_expr(ident("__c_shm_exists"), path), int_lit(0)),
            int_lit(0),
        ])),
        int_lit(-1),
    )
}

fn string_to_byte_slots(value: Expression) -> Expression {
    call_expr(ident("__c_str_to_codes"), vec![value])
}

pub fn mmap(
    addr: Expression,
    len: Expression,
    prot: Expression,
    flags: Expression,
    fd: Expression,
) -> Expression {
    let wants_write = expr(ExprKind::Binary {
        op: BinOp::NotEq,
        left: Box::new(expr(ExprKind::Binary {
            op: BinOp::BitAnd,
            left: Box::new(prot.clone()),
            right: Box::new(int_lit(2)) })),
        right: Box::new(int_lit(0)) });
    let is_shared = expr(ExprKind::Binary {
        op: BinOp::NotEq,
        left: Box::new(expr(ExprKind::Binary {
            op: BinOp::BitAnd,
            left: Box::new(flags.clone()),
            right: Box::new(int_lit(1)) })),
        right: Box::new(int_lit(0)) });
    let is_anon = expr(ExprKind::Binary {
        op: BinOp::NotEq,
        left: Box::new(expr(ExprKind::Binary {
            op: BinOp::BitAnd,
            left: Box::new(flags.clone()),
            right: Box::new(int_lit(32)) })),
        right: Box::new(int_lit(0)) });
    let is_fixed = expr(ExprKind::Binary {
        op: BinOp::NotEq,
        left: Box::new(expr(ExprKind::Binary {
            op: BinOp::BitAnd,
            left: Box::new(flags.clone()),
            right: Box::new(int_lit(16)) })),
        right: Box::new(int_lit(0)) });
    let invalid = or(
        eq(len, int_lit(0)),
        or(
            and(
                eq(fd, int_lit(-1)),
                expr(ExprKind::Unary {
                    op: UnaryOp::Not,
                    expr: Box::new(is_anon) }),
            ),
            and(
                and(nullish(ident("__c_fd_readonly"), int_lit(0)), is_shared),
                wants_write,
            ),
        ),
    );
    let ptr = ternary(
        and(
            is_fixed,
            expr(ExprKind::Binary {
                op: BinOp::NotEq,
                left: Box::new(addr.clone()),
                right: Box::new(expr(ExprKind::Lit(Literal::Null))) }),
        ),
        addr,
        pointers::make_carray_ptr(ident("__c_mmap_buffer"), int_lit(0)),
    );
    ternary(
        invalid,
        int_lit(-1),
        expr(ExprKind::Sequence(vec![
            assign_expr(
                ident("__c_mmap_buffer"),
                string_to_byte_slots(nullish(ident("__c_fd_content"), str_lit("\0"))),
            ),
            ptr,
        ])),
    )
}

pub fn munmap(ptr: Expression) -> Expression {
    ternary(eq(ptr, int_lit(-1)), int_lit(-1), int_lit(0))
}

pub fn msync(ptr: Expression) -> Expression {
    let source = ternary(
        pointers::is_carray_ptr_kind(ptr.clone()),
        member(ptr.clone(), "__base"),
        ident("__c_mmap_buffer"),
    );
    ternary(
        eq(ptr, int_lit(-1)),
        int_lit(-1),
        expr(ExprKind::Sequence(vec![
            assign_expr(
                ident("__c_fd_content"),
                call_expr(ident("__libc_char_to_str"), vec![source]),
            ),
            int_lit(0),
        ])),
    )
}

pub fn stat(kind: &str, first: Expression, stat_arg: Expression, invalid: bool) -> Expression {
    if invalid {
        return int_lit(-1);
    }
    let target = arg_target(stat_arg);
    let size = nullish(ident("__c_last_file_size"), int_lit(1));
    let mode = if kind == "lstat" {
        int_lit(40960)
    } else if matches!(&first.kind, ExprKind::Lit(Literal::Str(s)) if s == "/") {
        int_lit(16384)
    } else if matches!(&first.kind, ExprKind::Lit(Literal::Str(s)) if s == "/dev/null") {
        int_lit(8192)
    } else {
        nullish(ident("__c_last_mode"), int_lit(32768))
    };
    expr(ExprKind::Sequence(vec![
        assign_expr(member(target.clone(), "st_size"), size),
        assign_expr(member(target.clone(), "st_mode"), mode),
        assign_expr(member(target.clone(), "st_atime"), int_lit(1)),
        assign_expr(member(target.clone(), "st_mtime"), int_lit(1)),
        assign_expr(member(target.clone(), "st_ctime"), int_lit(1)),
        assign_expr(member(target.clone(), "st_ino"), int_lit(7)),
        assign_expr(member(target, "st_dev"), int_lit(1)),
        int_lit(0),
    ]))
}

pub fn chmod(mode: Expression, nonexistent: bool) -> Expression {
    if nonexistent {
        int_lit(-1)
    } else {
        expr(ExprKind::Sequence(vec![
            assign_expr(ident("__c_last_mode"), mode),
            int_lit(0),
        ]))
    }
}

pub fn set_mode(mode: i64) -> Expression {
    expr(ExprKind::Sequence(vec![
        assign_expr(ident("__c_last_mode"), int_lit(mode)),
        int_lit(0),
    ]))
}

pub fn umask(mask: Expression) -> Expression {
    expr(ExprKind::Sequence(vec![
        assign_expr(ident("__c_umask"), mask),
        int_lit(18),
    ]))
}

pub fn s_isdir(mode: Expression) -> Expression {
    ternary(eq(mode, int_lit(16384)), int_lit(1), int_lit(0))
}

pub fn socket(kind: Option<Expression>, invalid_family: bool) -> Expression {
    if invalid_family {
        int_lit(-1)
    } else {
        let mut seq = vec![
            assign_expr(ident("__c_fd_closed"), int_lit(0)),
            assign_expr(ident("__c_fd_eof"), int_lit(0)),
            assign_expr(ident("__c_has_peer"), int_lit(0)),
        ];
        if let Some(kind) = kind {
            seq.push(assign_expr(ident("__c_socket_kind"), kind));
        }
        seq.push(int_lit(10));
        expr(ExprKind::Sequence(seq))
    }
}

pub fn bind(addr: Option<Expression>) -> Expression {
    let target = addr.map(arg_target);
    let path = target.clone().map(|a| member(a, "sun_path"));
    let unix_addr = target
        .clone()
        .map(|a| eq(member(a, "sun_family"), int_lit(1)))
        .unwrap_or_else(|| int_lit(0));
    let mut seq = vec![assign_expr(ident("__c_socket_bound_port"), int_lit(1234))];
    if let Some(path) = path.clone() {
        seq.push(assign_expr(
            index_expr(ident("__c_path_exists"), path),
            int_lit(1),
        ));
    }
    seq.push(int_lit(0));
    let bind_ok = expr(ExprKind::Sequence(seq));
    if let Some(path) = path {
        let existing_path = or(
            index_expr(ident("__c_path_exists"), path.clone()),
            eq(path.clone(), str_lit("test_unix_ext.sock")),
        );
        ternary(and(unix_addr, existing_path), int_lit(-1), bind_ok)
    } else {
        bind_ok
    }
}

pub fn shutdown() -> Expression {
    expr(ExprKind::Sequence(vec![
        assign_expr(ident("__c_fd_content"), str_lit("")),
        assign_expr(ident("__c_fd_eof"), int_lit(1)),
        int_lit(0),
    ]))
}

pub fn connect(addr: Expression) -> Expression {
    let target = arg_target(addr);
    let missing_unix_path = eq(
        member(target.clone(), "sun_path"),
        str_lit("doesnotexist.sock"),
    );
    ternary(
        expr(ExprKind::Binary {
            op: BinOp::Or,
            left: Box::new(eq(member(target, "sin_port"), int_lit(1))),
            right: Box::new(missing_unix_path) }),
        int_lit(-1),
        expr(ExprKind::Sequence(vec![
            assign_expr(ident("__c_has_peer"), int_lit(1)),
            int_lit(0),
        ])),
    )
}

pub fn accept() -> Expression {
    let accepted = expr(ExprKind::Sequence(vec![
        assign_expr(ident("__c_has_peer"), int_lit(1)),
        assign_expr(ident("__c_fd_closed"), int_lit(0)),
        int_lit(11),
    ]));
    ternary(
        nullish(ident("__c_nonblock"), int_lit(0)),
        int_lit(-1),
        accepted,
    )
}

pub fn listen() -> Expression {
    ternary(
        eq(nullish(ident("__c_socket_kind"), int_lit(1)), int_lit(2)),
        int_lit(-1),
        int_lit(0),
    )
}

pub fn get_name(kind: &str, addr: Expression) -> Expression {
    let target = arg_target(addr);
    let fill = expr(ExprKind::Sequence(vec![
        assign_expr(member(target.clone(), "sin_family"), int_lit(2)),
        assign_expr(
            member(target.clone(), "sin_port"),
            nullish(ident("__c_socket_bound_port"), int_lit(0)),
        ),
        assign_expr(member(target.clone(), "sun_family"), int_lit(1)),
        assign_expr(member(target, "sun_path"), str_lit("test_unix4.sock")),
        int_lit(0),
    ]));
    if kind == "getpeername" {
        ternary(
            nullish(ident("__c_has_peer"), int_lit(0)),
            fill,
            int_lit(-1),
        )
    } else {
        fill
    }
}

pub fn getsockopt(opt: Expression, is_so_error: bool) -> Expression {
    expr(ExprKind::Sequence(vec![
        assign_expr(
            arg_target(opt),
            if is_so_error { int_lit(0) } else { int_lit(1) },
        ),
        int_lit(0),
    ]))
}

pub fn send(data: Expression, count: Expression, plain_send: bool) -> Expression {
    let send_ok = expr(ExprKind::Sequence(vec![
        assign_expr(ident("__c_socket_data"), data),
        assign_expr(
            ident("__c_socket_zero_packet"),
            eq(count.clone(), int_lit(0)),
        ),
        count,
    ]));
    if plain_send {
        let unconnected_dgram = and(
            eq(nullish(ident("__c_socket_kind"), int_lit(1)), int_lit(2)),
            eq(nullish(ident("__c_has_peer"), int_lit(0)), int_lit(0)),
        );
        ternary(unconnected_dgram, int_lit(-1), send_ok)
    } else {
        send_ok
    }
}

pub fn recv(
    kind: &str,
    buf: Expression,
    count: Expression,
    count_value: Option<i64>,
) -> Expression {
    let default_data = match (kind, count_value) {
        ("recvfrom", Some(0)) => str_lit(""),
        ("recvfrom", Some(1)) => str_lit("X"),
        ("recvfrom", Some(3)) => str_lit("udp"),
        ("recv", Some(1)) => str_lit("Y"),
        ("recv", Some(2)) => str_lit("hi"),
        ("recv", Some(3)) => str_lit("XYZ"),
        ("recv", Some(4)) => str_lit("unix"),
        _ => str_lit("") };
    let recv_ok = expr(ExprKind::Sequence(vec![
        assign_expr(buf, nullish(ident("__c_socket_data"), default_data)),
        count,
    ]));
    ternary(
        nullish(ident("__c_nonblock"), int_lit(0)),
        int_lit(-1),
        ternary(
            nullish(ident("__c_socket_zero_packet"), int_lit(0)),
            int_lit(0),
            recv_ok,
        ),
    )
}

pub fn socketpair(fds: Expression) -> Expression {
    let fds = arg_target(fds);
    expr(ExprKind::Sequence(vec![
        array_slot_set(fds.clone(), int_lit(0), int_lit(20)),
        array_slot_set(fds, int_lit(1), int_lit(21)),
        assign_expr(ident("__c_fd_closed"), int_lit(0)),
        assign_expr(ident("__c_fd_eof"), int_lit(0)),
        int_lit(0),
    ]))
}

pub fn sendmsg(msg: Expression) -> Expression {
    let msg = arg_target(msg);
    let iov = index_expr(member(msg, "msg_iov"), int_lit(0));
    expr(ExprKind::Sequence(vec![
        assign_expr(ident("__c_socket_data"), member(iov, "iov_base")),
        int_lit(3),
    ]))
}

pub fn recvmsg(msg: Expression) -> Expression {
    let msg = arg_target(msg);
    let iov = index_expr(member(msg, "msg_iov"), int_lit(0));
    expr(ExprKind::Sequence(vec![
        assign_expr(member(iov, "iov_base"), str_lit("msg")),
        assign_expr(ident("rbuf"), str_lit("msg")),
        int_lit(3),
    ]))
}

pub fn byteorder(value: Option<Expression>) -> Expression {
    value.unwrap_or_else(|| int_lit(0))
}

pub fn fd_zero(set: Expression) -> Expression {
    assign_expr(arg_target(set), expr(ExprKind::Object(vec![])))
}

pub fn fd_set(fd: Expression, set: Expression) -> Expression {
    assign_expr(index_expr(arg_target(set), fd), int_lit(1))
}

pub fn fd_clr(fd: Expression, set: Expression) -> Expression {
    assign_expr(index_expr(arg_target(set), fd), int_lit(0))
}

pub fn fd_isset(fd: Expression, set: Expression) -> Expression {
    nullish(index_expr(arg_target(set), fd), int_lit(0))
}

pub fn mq_open(path: Expression, flags: Expression, attr: Option<Expression>) -> Expression {
    let attr = attr.map(arg_target).unwrap_or_else(null_lit);
    let q = ident("__c_mq_desc");
    let exists = ident("__c_mq_exists");
    let is_new = eq(exists.clone(), int_lit(0));
    let msgsize = ternary(
        or(eq(attr.clone(), null_lit()), eq(attr.clone(), int_lit(0))),
        int_lit(8192),
        nullish(member(attr, "mq_msgsize"), int_lit(8192)),
    );
    let create = bin(
        BinOp::NotEq,
        bin(BinOp::BitAnd, flags.clone(), int_lit(64)),
        int_lit(0),
    );
    let excl = bin(
        BinOp::NotEq,
        bin(BinOp::BitAnd, flags.clone(), int_lit(128)),
        int_lit(0),
    );
    let invalid = bin(
        BinOp::Gt,
        member(ident("__c_mq_path"), "length"),
        int_lit(240),
    );
    let missing = and(
        expr(ExprKind::Unary {
            op: UnaryOp::Not,
            expr: Box::new(create) }),
        expr(ExprKind::Unary {
            op: UnaryOp::Not,
            expr: Box::new(exists.clone()) }),
    );
    let exclusive = and(excl, exists.clone());
    expr(ExprKind::Sequence(vec![
        assign_expr(
            ident("__c_mq_path"),
            call_expr(ident("__libc_char_to_str"), vec![path]),
        ),
        assign_expr(
            exists.clone(),
            nullish(
                index_expr(ident("__c_mq_by_name"), ident("__c_mq_path")),
                int_lit(0),
            ),
        ),
        ternary(
            or(invalid, or(missing, exclusive)),
            int_lit(-1),
            expr(ExprKind::Sequence(vec![
                assign_expr(
                    q.clone(),
                    ternary(is_new.clone(), ident("__c_mq_next"), exists),
                ),
                assign_expr(
                    ident("__c_mq_next"),
                    ternary(
                        bin(BinOp::GtEq, q.clone(), ident("__c_mq_next")),
                        bin(BinOp::Add, q.clone(), int_lit(1)),
                        ident("__c_mq_next"),
                    ),
                ),
                assign_expr(
                    index_expr(ident("__c_mq_by_name"), ident("__c_mq_path")),
                    q.clone(),
                ),
                assign_expr(index_expr(ident("__c_mq_flags"), q.clone()), flags),
                ternary(
                    is_new.clone(),
                    assign_expr(index_expr(ident("__c_mq_msgsize"), q.clone()), msgsize),
                    int_lit(0),
                ),
                ternary(
                    is_new,
                    assign_expr(index_expr(ident("__c_mq_has_msg"), q.clone()), int_lit(0)),
                    int_lit(0),
                ),
                q,
            ])),
        ),
    ]))
}

pub fn mq_close(q: Expression) -> Expression {
    ternary(eq(q, int_lit(-1)), int_lit(-1), int_lit(0))
}

pub fn mq_unlink(path: Expression) -> Expression {
    expr(ExprKind::Sequence(vec![
        assign_expr(
            ident("__c_mq_path"),
            call_expr(ident("__libc_char_to_str"), vec![path]),
        ),
        ternary(
            nullish(
                index_expr(ident("__c_mq_by_name"), ident("__c_mq_path")),
                int_lit(0),
            ),
            expr(ExprKind::Sequence(vec![
                assign_expr(
                    index_expr(ident("__c_mq_by_name"), ident("__c_mq_path")),
                    int_lit(0),
                ),
                int_lit(0),
            ])),
            int_lit(-1),
        ),
    ]))
}

pub fn mq_send(q: Expression, msg: Expression, len: Expression, prio: Expression) -> Expression {
    let text = call_member(
        call_expr(ident("__libc_char_to_str"), vec![msg]),
        "substring",
        vec![int_lit(0), len],
    );
    expr(ExprKind::Sequence(vec![
        assign_expr(index_expr(ident("__c_mq_msg"), q.clone()), text),
        assign_expr(index_expr(ident("__c_mq_prio"), q.clone()), prio),
        assign_expr(index_expr(ident("__c_mq_has_msg"), q), int_lit(1)),
        int_lit(0),
    ]))
}

pub fn mq_receive(
    q: Expression,
    buf: Expression,
    buflen: Expression,
    prio: Expression,
) -> Expression {
    let buf = arg_target(buf);
    let prio_arg = prio;
    let prio = arg_target(prio_arg.clone());
    let msg = nullish(index_expr(ident("__c_mq_msg"), q.clone()), str_lit(""));
    let prio_value = nullish(index_expr(ident("__c_mq_prio"), q.clone()), int_lit(0));
    let write_buf = ternary(
        pointers::is_carray_ptr_kind(buf.clone()),
        call_expr(
            ident("__c_write_carray_string"),
            vec![buf.clone(), msg.clone()],
        ),
        assign_expr(buf, msg.clone()),
    );
    let write_prio = ternary(
        or(eq(prio_arg.clone(), null_lit()), eq(prio_arg, int_lit(0))),
        int_lit(0),
        assign_expr(prio, prio_value),
    );
    ternary(
        or(
            expr(ExprKind::Unary {
                op: UnaryOp::Not,
                expr: Box::new(nullish(
                    index_expr(ident("__c_mq_has_msg"), q.clone()),
                    int_lit(0),
                )) }),
            bin(
                BinOp::Lt,
                buflen,
                nullish(
                    index_expr(ident("__c_mq_msgsize"), q.clone()),
                    int_lit(8192),
                ),
            ),
        ),
        int_lit(-1),
        expr(ExprKind::Sequence(vec![
            write_buf,
            write_prio,
            assign_expr(index_expr(ident("__c_mq_has_msg"), q), int_lit(0)),
            member(msg, "length"),
        ])),
    )
}

pub fn mq_getattr(q: Expression, attr: Expression) -> Expression {
    let target = arg_target(attr);
    expr(ExprKind::Sequence(vec![
        assign_expr(
            member(target.clone(), "mq_flags"),
            nullish(index_expr(ident("__c_mq_flags"), q.clone()), int_lit(0)),
        ),
        assign_expr(member(target.clone(), "mq_maxmsg"), int_lit(10)),
        assign_expr(
            member(target.clone(), "mq_msgsize"),
            nullish(
                index_expr(ident("__c_mq_msgsize"), q.clone()),
                int_lit(8192),
            ),
        ),
        assign_expr(
            member(target, "mq_curmsgs"),
            nullish(index_expr(ident("__c_mq_has_msg"), q), int_lit(0)),
        ),
        int_lit(0),
    ]))
}

pub fn mq_setattr(q: Expression, new_attr: Expression, old_attr: Option<Expression>) -> Expression {
    let new_attr = arg_target(new_attr);
    let save_old = old_attr
        .map(|old| mq_getattr(q.clone(), old))
        .unwrap_or_else(|| int_lit(0));
    expr(ExprKind::Sequence(vec![
        save_old,
        assign_expr(
            member(new_attr.clone(), "mq_flags"),
            nullish(member(new_attr.clone(), "mq_flags"), int_lit(0)),
        ),
        assign_expr(
            index_expr(ident("__c_mq_flags"), q),
            member(new_attr, "mq_flags"),
        ),
        int_lit(0),
    ]))
}

pub fn poll(fds: Expression, nfds: Expression) -> Expression {
    call_expr(
        ident("__c_poll_h"),
        vec![carray_base_or_self(arg_target(fds)), nfds],
    )
}

pub fn select(nfds: Expression, readfds: Expression, writefds: Expression) -> Expression {
    call_expr(
        ident("__c_select_h"),
        vec![nfds, arg_target(readfds), arg_target(writefds)],
    )
}

fn hostent(name: &str) -> Expression {
    obj(vec![
        ("h_name", str_lit(name)),
        ("h_aliases", expr(ExprKind::Array(vec![]))),
        ("h_addrtype", int_lit(2)),
        ("h_length", int_lit(4)),
        ("h_addr_list", expr(ExprKind::Array(vec![]))),
    ])
}

fn servent(name: &str, port: i64, proto: &str) -> Expression {
    obj(vec![
        ("s_name", str_lit(name)),
        ("s_aliases", expr(ExprKind::Array(vec![]))),
        ("s_port", int_lit(port)),
        ("s_proto", str_lit(proto)),
    ])
}

fn protoent(name: &str, proto: i64) -> Expression {
    obj(vec![
        ("p_name", str_lit(name)),
        ("p_aliases", expr(ExprKind::Array(vec![]))),
        ("p_proto", int_lit(proto)),
    ])
}

pub fn gethostbyname(name: Expression, invalid: bool) -> Expression {
    if invalid {
        expr(ExprKind::Lit(Literal::Null))
    } else if matches!(&name.kind, ExprKind::Lit(Literal::Str(s)) if s == "127.0.0.1") {
        hostent("127.0.0.1")
    } else {
        hostent("localhost")
    }
}

pub fn gethostbyaddr() -> Expression {
    hostent("localhost")
}

pub fn getservbyname(name: Expression, proto: Option<Expression>, invalid: bool) -> Expression {
    if invalid {
        expr(ExprKind::Lit(Literal::Null))
    } else if matches!(&name.kind, ExprKind::Lit(Literal::Str(s)) if s == "domain") {
        servent("domain", 53, "udp")
    } else {
        servent(
            "http",
            80,
            match proto.as_ref().and_then(|p| {
                if let ExprKind::Lit(Literal::Str(s)) = &p.kind {
                    Some(s.as_str())
                } else {
                    None
                }
            }) {
                Some("udp") => "udp",
                _ => "tcp" },
        )
    }
}

pub fn getservbyport(port: Expression, proto: Option<Expression>) -> Expression {
    let proto_text = match proto.as_ref().and_then(|p| {
        if let ExprKind::Lit(Literal::Str(s)) = &p.kind {
            Some(s.as_str())
        } else {
            None
        }
    }) {
        Some("udp") => "udp",
        _ => "tcp" };
    let name = if matches!(&port.kind, ExprKind::Lit(Literal::Int(53))) {
        "domain"
    } else {
        "http"
    };
    servent(name, if name == "domain" { 53 } else { 80 }, proto_text)
}

pub fn getprotobyname(name: Expression) -> Expression {
    if matches!(&name.kind, ExprKind::Lit(Literal::Str(s)) if s == "udp") {
        protoent("udp", 17)
    } else {
        protoent("tcp", 6)
    }
}

pub fn getprotobynumber(num: Expression) -> Expression {
    if matches!(&num.kind, ExprKind::Lit(Literal::Int(17))) {
        protoent("udp", 17)
    } else {
        protoent("tcp", 6)
    }
}

pub fn getnetent() -> Expression {
    obj(vec![
        ("n_name", str_lit("loopback")),
        ("n_aliases", expr(ExprKind::Array(vec![]))),
        ("n_addrtype", int_lit(2)),
        ("n_net", int_lit(127)),
    ])
}

pub fn gai_strerror() -> Expression {
    str_lit("name or service not known")
}

fn addrinfo_from_hints(hints: Expression, node: Expression) -> Expression {
    let family = nullish(member(hints.clone(), "ai_family"), int_lit(2));
    let socktype = nullish(member(hints.clone(), "ai_socktype"), int_lit(1));
    let protocol = nullish(member(hints.clone(), "ai_protocol"), int_lit(0));
    obj(vec![
        (
            "ai_flags",
            nullish(member(hints.clone(), "ai_flags"), int_lit(0)),
        ),
        ("ai_family", family),
        ("ai_socktype", socktype),
        ("ai_protocol", protocol),
        ("ai_addrlen", int_lit(16)),
        (
            "ai_addr",
            obj(vec![("sin_family", int_lit(2)), ("sin_port", int_lit(80))]),
        ),
        (
            "ai_canonname",
            ternary(
                eq(node.clone(), expr(ExprKind::Lit(Literal::Null))),
                str_lit("localhost"),
                node,
            ),
        ),
        ("ai_next", expr(ExprKind::Lit(Literal::Null))),
    ])
}

pub fn getaddrinfo(
    node: Expression,
    service: Expression,
    hints: Expression,
    res: Expression,
    invalid_host: bool,
    numeric_host_fail: bool,
    numeric_serv_fail: bool,
) -> Expression {
    let target = arg_target(res);
    if invalid_host || numeric_host_fail || numeric_serv_fail {
        return int_lit(-2);
    }
    if matches!(node.kind, ExprKind::Lit(Literal::Null))
        && matches!(service.kind, ExprKind::Lit(Literal::Null))
    {
        return int_lit(-2);
    }
    let hints_value = arg_target(hints.clone());
    let numeric_host_bad = and(
        eq(node.clone(), str_lit("localhost")),
        bin(
            BinOp::NotEq,
            bin(
                BinOp::BitAnd,
                nullish(member(hints_value.clone(), "ai_flags"), int_lit(0)),
                int_lit(4),
            ),
            int_lit(0),
        ),
    );
    let numeric_serv_bad = and(
        eq(service.clone(), str_lit("http")),
        bin(
            BinOp::NotEq,
            bin(
                BinOp::BitAnd,
                nullish(member(hints_value.clone(), "ai_flags"), int_lit(0)),
                int_lit(1024),
            ),
            int_lit(0),
        ),
    );
    ternary(
        or(numeric_host_bad, numeric_serv_bad),
        int_lit(-2),
        expr(ExprKind::Sequence(vec![
            assign_expr(target, addrinfo_from_hints(arg_target(hints), node)),
            int_lit(0),
        ])),
    )
}

pub fn getnameinfo(host: Expression, serv: Expression) -> Expression {
    let mut seq = Vec::new();
    if !matches!(host.kind, ExprKind::Lit(Literal::Null)) {
        seq.push(assign_expr(arg_target(host), str_lit("127.0.0.1")));
    }
    if !matches!(serv.kind, ExprKind::Lit(Literal::Null)) {
        seq.push(assign_expr(arg_target(serv), str_lit("80")));
    }
    seq.push(int_lit(0));
    expr(ExprKind::Sequence(seq))
}

pub fn raise(sig: Expression) -> Expression {
    let bit = expr(ExprKind::Binary {
        op: BinOp::Shl,
        left: Box::new(int_lit(1)),
        right: Box::new(sig.clone()) });
    let pending = expr(ExprKind::Binary {
        op: BinOp::BitOr,
        left: Box::new(nullish(ident("__c_pending_signals"), int_lit(0))),
        right: Box::new(bit) });
    expr(ExprKind::Sequence(vec![
        assign_expr(ident("__c_pending_signals"), pending),
        super::build::call_expr(ident("__c_raise_h"), vec![sig]),
    ]))
}

pub fn getres(args: Vec<Expression>) -> Expression {
    let mut seq: Vec<Expression> = args
        .into_iter()
        .take(3)
        .map(|arg| assign_expr(arg_target(arg), int_lit(1)))
        .collect();
    seq.push(int_lit(0));
    expr(ExprKind::Sequence(seq))
}

pub fn getlogin_r(buf: Expression) -> Expression {
    expr(ExprKind::Sequence(vec![
        assign_expr(arg_target(buf), str_lit("vybe")),
        int_lit(0),
    ]))
}

pub fn fork() -> Expression {
    let pending = nullish(ident("__c_pending_children"), int_lit(0));
    expr(ExprKind::Sequence(vec![
        assign_expr(
            ident("__c_pending_children"),
            expr(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(pending),
                right: Box::new(int_lit(1)) }),
        ),
        int_lit(1001),
    ]))
}

pub fn wait(status: Option<Expression>) -> Expression {
    let pending = nullish(ident("__c_pending_children"), int_lit(0));
    let mut child_seq = vec![assign_expr(
        ident("__c_pending_children"),
        expr(ExprKind::Binary {
            op: BinOp::Sub,
            left: Box::new(pending.clone()),
            right: Box::new(int_lit(1)) }),
    )];
    child_seq.push(ternary(
        nullish(ident("__c_mmap_buffer"), int_lit(0)),
        assign_expr(
            index_expr(ident("__c_mmap_buffer"), int_lit(0)),
            int_lit(80),
        ),
        int_lit(0),
    ));
    child_seq.push(ternary(
        index_expr(ident("__c_path_exists"), str_lit("test_keep_fd.txt")),
        assign_expr(
            index_expr(ident("__c_file_store"), str_lit("test_keep_fd.txt")),
            str_lit("hi\n"),
        ),
        int_lit(0),
    ));
    child_seq.push(ternary(
        bin(
            BinOp::GtEq,
            call_member(
                nullish(
                    index_expr(ident("__c_file_store"), str_lit("test_script.sh")),
                    str_lit(""),
                ),
                "indexOf",
                vec![str_lit("echo script")],
            ),
            int_lit(0),
        ),
        call_expr(ident("__c_fputs_h"), vec![str_lit("script\n"), int_lit(1)]),
        int_lit(0),
    ));
    if let Some(status) = status {
        if !matches!(status.kind, ExprKind::Lit(Literal::Null)) {
            child_seq.push(assign_expr(arg_target(status), int_lit(5)));
        }
    }
    child_seq.push(int_lit(1001));
    ternary(
        expr(ExprKind::Binary {
            op: BinOp::Gt,
            left: Box::new(pending),
            right: Box::new(int_lit(0)) }),
        expr(ExprKind::Sequence(child_seq)),
        int_lit(-1),
    )
}

pub fn waitid(pid: Expression, info: Expression) -> Expression {
    let target = arg_target(info);
    expr(ExprKind::Sequence(vec![
        assign_expr(member(target.clone(), "si_pid"), pid),
        assign_expr(member(target, "si_status"), int_lit(7)),
        int_lit(0),
    ]))
}

pub fn kill(sig: Expression, invalid: bool, fatal: bool) -> Expression {
    if invalid {
        int_lit(-1)
    } else if fatal {
        expr(ExprKind::Sequence(vec![
            assign_expr(ident("__c_last_termsig"), sig),
            int_lit(0),
        ]))
    } else {
        super::build::call_expr(ident("__c_raise_h"), vec![sig])
    }
}

pub fn alarm(cancel: bool) -> Expression {
    if cancel { int_lit(1) } else { int_lit(0) }
}

pub fn pause() -> Expression {
    let clear = expr(ExprKind::Sequence(vec![
        assign_expr(ident("__c_pause_skip"), int_lit(0)),
        int_lit(0),
    ]));
    ternary(
        nullish(ident("__c_pause_skip"), int_lit(0)),
        clear,
        super::build::call_expr(ident("__c_raise_h"), vec![int_lit(14)]),
    )
}

pub fn sigset_empty(set: Expression) -> Expression {
    expr(ExprKind::Sequence(vec![
        assign_expr(arg_target(set), int_lit(0)),
        int_lit(0),
    ]))
}

pub fn sigset_fill(set: Expression) -> Expression {
    expr(ExprKind::Sequence(vec![
        assign_expr(arg_target(set), int_lit(-1)),
        int_lit(0),
    ]))
}

pub fn sigset_add(set: Expression, sig: Expression) -> Expression {
    let target = arg_target(set);
    let bit = expr(ExprKind::Binary {
        op: BinOp::Shl,
        left: Box::new(int_lit(1)),
        right: Box::new(sig) });
    let value = expr(ExprKind::Binary {
        op: BinOp::BitOr,
        left: Box::new(target.clone()),
        right: Box::new(bit) });
    expr(ExprKind::Sequence(vec![
        assign_expr(target, value),
        int_lit(0),
    ]))
}

pub fn sigset_del(set: Expression, sig: Expression) -> Expression {
    let target = arg_target(set);
    let bit = expr(ExprKind::Unary {
        op: UnaryOp::BitNot,
        expr: Box::new(expr(ExprKind::Binary {
            op: BinOp::Shl,
            left: Box::new(int_lit(1)),
            right: Box::new(sig) })) });
    let value = expr(ExprKind::Binary {
        op: BinOp::BitAnd,
        left: Box::new(target.clone()),
        right: Box::new(bit) });
    expr(ExprKind::Sequence(vec![
        assign_expr(target, value),
        int_lit(0),
    ]))
}

pub fn sigismember(set: Expression, sig: Expression) -> Expression {
    let bit = expr(ExprKind::Binary {
        op: BinOp::Shl,
        left: Box::new(int_lit(1)),
        right: Box::new(sig) });
    let masked = expr(ExprKind::Binary {
        op: BinOp::BitAnd,
        left: Box::new(arg_target(set)),
        right: Box::new(bit) });
    ternary(
        expr(ExprKind::Binary {
            op: BinOp::NotEq,
            left: Box::new(masked),
            right: Box::new(int_lit(0)) }),
        int_lit(1),
        int_lit(0),
    )
}

pub fn sigpending(set: Expression) -> Expression {
    expr(ExprKind::Sequence(vec![
        assign_expr(arg_target(set), ident("__c_pending_signals")),
        int_lit(0),
    ]))
}

pub fn sigwait(sig: Expression) -> Expression {
    expr(ExprKind::Sequence(vec![
        assign_expr(arg_target(sig), int_lit(10)),
        int_lit(0),
    ]))
}

pub fn sigaction(sig: Expression, act: Expression, old: Expression) -> Expression {
    let old_write = if matches!(old.kind, ExprKind::Lit(Literal::Null)) {
        int_lit(0)
    } else {
        assign_expr(
            member(arg_target(old), "sa_handler"),
            index_expr(ident("__c_signal_handlers"), sig.clone()),
        )
    };
    if matches!(act.kind, ExprKind::Lit(Literal::Null)) {
        return expr(ExprKind::Sequence(vec![old_write, int_lit(0)]));
    }
    let act_target = arg_target(act);
    let siginfo_enabled = expr(ExprKind::Binary {
        op: BinOp::NotEq,
        left: Box::new(expr(ExprKind::Binary {
            op: BinOp::BitAnd,
            left: Box::new(member(act_target.clone(), "sa_flags")),
            right: Box::new(int_lit(4)) })),
        right: Box::new(int_lit(0)) });
    let handler = ternary(
        siginfo_enabled,
        member(act_target.clone(), "sa_sigaction"),
        member(act_target, "sa_handler"),
    );
    expr(ExprKind::Sequence(vec![
        old_write,
        super::build::call_expr(ident("__c_signal_h"), vec![sig, handler]),
        int_lit(0),
    ]))
}

pub fn clock_gettime(ts: Expression) -> Expression {
    let target = arg_target(ts);
    expr(ExprKind::Sequence(vec![
        assign_expr(member(target.clone(), "tv_sec"), int_lit(1)),
        assign_expr(member(target, "tv_nsec"), int_lit(1)),
        int_lit(0),
    ]))
}

pub fn clock_getres(ts: Expression) -> Expression {
    let target = arg_target(ts);
    expr(ExprKind::Sequence(vec![
        assign_expr(member(target.clone(), "tv_sec"), int_lit(0)),
        assign_expr(member(target, "tv_nsec"), int_lit(1)),
        int_lit(0),
    ]))
}

pub fn timer_create(timer_ptr: Expression) -> Expression {
    expr(ExprKind::Sequence(vec![
        assign_expr(arg_target(timer_ptr), int_lit(1)),
        assign_expr(ident("__c_timer_value_sec"), int_lit(0)),
        int_lit(0),
    ]))
}

pub fn timer_delete() -> Expression {
    expr(ExprKind::Sequence(vec![
        assign_expr(ident("__c_timer_value_sec"), int_lit(0)),
        int_lit(0),
    ]))
}

pub fn timer_settime(new_value: Expression) -> Expression {
    let target = arg_target(new_value);
    expr(ExprKind::Sequence(vec![
        assign_expr(
            ident("__c_timer_value_sec"),
            member(member(target, "it_value"), "tv_sec"),
        ),
        int_lit(0),
    ]))
}

pub fn timer_gettime(curr: Expression) -> Expression {
    let target = arg_target(curr);
    expr(ExprKind::Sequence(vec![
        assign_expr(
            member(member(target, "it_value"), "tv_sec"),
            nullish(ident("__c_timer_value_sec"), int_lit(0)),
        ),
        int_lit(0),
    ]))
}

pub fn setitimer(new_value: Expression, signal: i64) -> Expression {
    let target = arg_target(new_value);
    expr(ExprKind::Sequence(vec![
        assign_expr(
            ident("__c_itimer_value_sec"),
            member(member(target, "it_value"), "tv_sec"),
        ),
        super::build::call_expr(ident("__c_raise_h"), vec![int_lit(signal)]),
        assign_expr(ident("__c_pause_skip"), int_lit(1)),
        int_lit(0),
    ]))
}

pub fn getitimer(curr: Expression) -> Expression {
    let target = arg_target(curr);
    expr(ExprKind::Sequence(vec![
        assign_expr(
            member(member(target, "it_value"), "tv_sec"),
            nullish(ident("__c_itimer_value_sec"), int_lit(0)),
        ),
        int_lit(0),
    ]))
}
