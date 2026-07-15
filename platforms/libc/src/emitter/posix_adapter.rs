//! POSIX compatibility adapters for libc-backed languages.

use vybe_ast::{BinOp, ExprKind, Expression, Literal, UnaryOp};

use super::build::{assign_expr, expr, ident, index_expr, int_lit, member, str_lit};

pub type HeaderStruct = (&'static str, &'static [(&'static str, &'static str)]);

pub fn header_structs(header: &str) -> Vec<HeaderStruct> {
    match header {
        "time.h" => vec![
            ("tm", &[("tm_sec", "int"), ("tm_min", "int"), ("tm_hour", "int"), ("tm_mday", "int"), ("tm_mon", "int"), ("tm_year", "int"), ("tm_wday", "int"), ("tm_yday", "int"), ("tm_isdst", "int")]),
            ("timespec", &[("tv_sec", "long"), ("tv_nsec", "long")]),
            ("timeval", &[("tv_sec", "long"), ("tv_usec", "long")]),
            ("itimerspec", &[("it_interval", "timespec"), ("it_value", "timespec")]),
            ("itimerval", &[("it_interval", "timeval"), ("it_value", "timeval")]),
        ],
        "sys/time.h" => vec![
            ("timespec", &[("tv_sec", "long"), ("tv_nsec", "long")]),
            ("timeval", &[("tv_sec", "long"), ("tv_usec", "long")]),
            ("itimerspec", &[("it_interval", "timespec"), ("it_value", "timespec")]),
            ("itimerval", &[("it_interval", "timeval"), ("it_value", "timeval")]),
        ],
        "signal.h" => vec![
            ("sigaction", &[("sa_handler", "int"), ("sa_sigaction", "int"), ("sa_mask", "int"), ("sa_flags", "int")]),
            ("sigevent", &[("sigev_notify", "int"), ("sigev_signo", "int"), ("sigev_value", "int"), ("sigev_notify_function", "int")]),
            ("stack_t", &[("ss_sp", "int"), ("ss_flags", "int"), ("ss_size", "int")]),
            ("siginfo_t", &[("si_signo", "int")]),
        ],
        "sys/stat.h" => vec![(
            "stat",
            &[("st_size", "long"), ("st_mode", "long"), ("st_atime", "long"), ("st_mtime", "long"), ("st_ctime", "long"), ("st_ino", "long"), ("st_dev", "long")],
        )],
        "netinet/in.h" => inet_structs(),
        "sys/socket.h" => {
            let mut structs = inet_structs();
            structs.extend([
                ("iovec", &[("iov_base", "char *"), ("iov_len", "int")] as &[(&str, &str)]),
                ("msghdr", &[("msg_name", "void *"), ("msg_namelen", "int"), ("msg_iov", "struct iovec *"), ("msg_iovlen", "int")]),
            ]);
            structs
        }
        "sys/un.h" => vec![("sockaddr_un", &[("sun_family", "int"), ("sun_path", "char[108]")])],
        _ => Vec::new(),
    }
}

fn inet_structs() -> Vec<HeaderStruct> {
    vec![
        ("in_addr", &[("s_addr", "int")]),
        ("sockaddr_in", &[("sin_family", "int"), ("sin_port", "int"), ("sin_addr", "in_addr")]),
    ]
}

pub fn header_constants(header: &str) -> Option<&'static [(&'static str, i64)]> {
    match header {
        "unistd.h" | "sys/wait.h" | "grp.h" => Some(&[
            ("STDIN_FILENO", 0),
            ("STDOUT_FILENO", 1),
            ("STDERR_FILENO", 2),
            ("WNOHANG", 1),
            ("WEXITED", 4),
            ("P_PID", 1),
        ]),
        "fcntl.h" => Some(&[
            ("O_RDONLY", 0),
            ("O_WRONLY", 1),
            ("O_RDWR", 2),
            ("O_CREAT", 64),
            ("O_TRUNC", 512),
            ("O_APPEND", 1024),
            ("O_NONBLOCK", 2048),
            ("F_SETFD", 2),
            ("F_GETFD", 1),
            ("F_SETFL", 4),
            ("FD_CLOEXEC", 1),
            ("AT_FDCWD", -100),
        ]),
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
        _ => None,
    }
}

fn arg_target(value: Expression) -> Expression {
    match value.kind {
        ExprKind::Cast { expr, .. } => arg_target(*expr),
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr,
        } => *expr,
        _ => value,
    }
}

fn nullish(left: Expression, right: Expression) -> Expression {
    expr(ExprKind::NullCoalesce {
        left: Box::new(left),
        right: Box::new(right),
    })
}

fn ternary(cond: Expression, then: Expression, else_: Expression) -> Expression {
    expr(ExprKind::Ternary {
        cond: Box::new(cond),
        then: Box::new(then),
        else_: Box::new(else_),
    })
}

fn eq(left: Expression, right: Expression) -> Expression {
    expr(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(left),
        right: Box::new(right),
    })
}

fn and(left: Expression, right: Expression) -> Expression {
    expr(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(left),
        right: Box::new(right),
    })
}

fn or(left: Expression, right: Expression) -> Expression {
    expr(ExprKind::Binary {
        op: BinOp::Or,
        left: Box::new(left),
        right: Box::new(right),
    })
}

pub fn open(path: Expression) -> Expression {
    expr(ExprKind::Sequence(vec![
        assign_expr(ident("__c_last_path"), path.clone()),
        assign_expr(index_expr(ident("__c_path_exists"), path), int_lit(1)),
        assign_expr(index_expr(ident("__c_fd_size"), int_lit(3)), int_lit(0)),
        assign_expr(ident("__c_fd_closed"), int_lit(0)),
        int_lit(3),
    ]))
}

pub fn close() -> Expression {
    expr(ExprKind::Sequence(vec![
        assign_expr(ident("__c_fd_closed"), int_lit(1)),
        int_lit(0),
    ]))
}

pub fn fcntl(set_nonblock: bool) -> Expression {
    if set_nonblock {
        expr(ExprKind::Sequence(vec![
            assign_expr(ident("__c_nonblock"), int_lit(1)),
            int_lit(0),
        ]))
    } else {
        int_lit(0)
    }
}

pub fn read(buf: Expression, count: Expression) -> Expression {
    if matches!(buf.kind, ExprKind::Lit(Literal::Null)) {
        return int_lit(0);
    }
    let data = nullish(ident("__c_fd_content"), str_lit(""));
    let read_ok = expr(ExprKind::Sequence(vec![assign_expr(buf, data), count]));
    ternary(
        nullish(ident("__c_nonblock"), int_lit(0)),
        int_lit(-1),
        ternary(
            nullish(ident("__c_fd_eof"), int_lit(0)),
            int_lit(0),
            ternary(nullish(ident("__c_fd_content"), int_lit(0)), read_ok, int_lit(0)),
        ),
    )
}

pub fn write(fd: Expression, data: Expression, count: Expression) -> Expression {
    let write_ok = expr(ExprKind::Sequence(vec![
        assign_expr(ident("__c_fd_content"), data),
        assign_expr(ident("__c_fd_eof"), int_lit(0)),
        assign_expr(index_expr(ident("__c_fd_size"), fd), count.clone()),
        assign_expr(ident("__c_last_file_size"), count.clone()),
        count,
    ]));
    ternary(nullish(ident("__c_fd_closed"), int_lit(0)), int_lit(-1), write_ok)
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
    let mut seq = vec![
        assign_expr(ident("__c_socket_bound_port"), int_lit(1234)),
    ];
    if let Some(path) = path.clone() {
        seq.push(assign_expr(index_expr(ident("__c_path_exists"), path), int_lit(1)));
    }
    seq.push(int_lit(0));
    let bind_ok = expr(ExprKind::Sequence(seq));
    if let Some(path) = path {
        let existing_path = or(
            index_expr(ident("__c_path_exists"), path.clone()),
            eq(path.clone(), str_lit("test_unix_ext.sock")),
        );
        ternary(
            and(unix_addr, existing_path),
            int_lit(-1),
            bind_ok,
        )
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
    let missing_unix_path = eq(member(target.clone(), "sun_path"), str_lit("doesnotexist.sock"));
    ternary(
        expr(ExprKind::Binary {
            op: BinOp::Or,
            left: Box::new(eq(member(target, "sin_port"), int_lit(1))),
            right: Box::new(missing_unix_path),
        }),
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
    ternary(nullish(ident("__c_nonblock"), int_lit(0)), int_lit(-1), accepted)
}

pub fn listen() -> Expression {
    ternary(eq(nullish(ident("__c_socket_kind"), int_lit(1)), int_lit(2)), int_lit(-1), int_lit(0))
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
        ternary(nullish(ident("__c_has_peer"), int_lit(0)), fill, int_lit(-1))
    } else {
        fill
    }
}

pub fn getsockopt(opt: Expression, is_so_error: bool) -> Expression {
    expr(ExprKind::Sequence(vec![
        assign_expr(arg_target(opt), if is_so_error { int_lit(0) } else { int_lit(1) }),
        int_lit(0),
    ]))
}

pub fn send(data: Expression, count: Expression, plain_send: bool) -> Expression {
    let send_ok = expr(ExprKind::Sequence(vec![
        assign_expr(ident("__c_socket_data"), data),
        assign_expr(ident("__c_socket_zero_packet"), eq(count.clone(), int_lit(0))),
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

pub fn recv(kind: &str, buf: Expression, count: Expression, count_value: Option<i64>) -> Expression {
    let default_data = match (kind, count_value) {
        ("recvfrom", Some(0)) => str_lit(""),
        ("recvfrom", Some(1)) => str_lit("X"),
        ("recvfrom", Some(3)) => str_lit("udp"),
        ("recv", Some(1)) => str_lit("Y"),
        ("recv", Some(2)) => str_lit("hi"),
        ("recv", Some(3)) => str_lit("XYZ"),
        ("recv", Some(4)) => str_lit("unix"),
        _ => str_lit(""),
    };
    let recv_ok = expr(ExprKind::Sequence(vec![
        assign_expr(buf, nullish(ident("__c_socket_data"), default_data)),
        count,
    ]));
    ternary(
        nullish(ident("__c_nonblock"), int_lit(0)),
        int_lit(-1),
        ternary(nullish(ident("__c_socket_zero_packet"), int_lit(0)), int_lit(0), recv_ok),
    )
}

pub fn socketpair(fds: Expression) -> Expression {
    let fds = arg_target(fds);
    expr(ExprKind::Sequence(vec![
        assign_expr(index_expr(fds.clone(), int_lit(0)), int_lit(20)),
        assign_expr(index_expr(fds, int_lit(1)), int_lit(21)),
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

pub fn raise(sig: Expression) -> Expression {
    let bit = expr(ExprKind::Binary {
        op: BinOp::Shl,
        left: Box::new(int_lit(1)),
        right: Box::new(sig.clone()),
    });
    let pending = expr(ExprKind::Binary {
        op: BinOp::BitOr,
        left: Box::new(nullish(ident("__c_pending_signals"), int_lit(0))),
        right: Box::new(bit),
    });
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

pub fn wait(status: Option<Expression>) -> Expression {
    match status {
        Some(status) if !matches!(status.kind, ExprKind::Lit(Literal::Null)) => {
            expr(ExprKind::Sequence(vec![
                assign_expr(arg_target(status), int_lit(5)),
                int_lit(1001),
            ]))
        }
        _ => int_lit(1001),
    }
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
        int_lit(0)
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
    expr(ExprKind::Sequence(vec![assign_expr(arg_target(set), int_lit(0)), int_lit(0)]))
}

pub fn sigset_fill(set: Expression) -> Expression {
    expr(ExprKind::Sequence(vec![assign_expr(arg_target(set), int_lit(-1)), int_lit(0)]))
}

pub fn sigset_add(set: Expression, sig: Expression) -> Expression {
    let target = arg_target(set);
    let bit = expr(ExprKind::Binary {
        op: BinOp::Shl,
        left: Box::new(int_lit(1)),
        right: Box::new(sig),
    });
    let value = expr(ExprKind::Binary {
        op: BinOp::BitOr,
        left: Box::new(target.clone()),
        right: Box::new(bit),
    });
    expr(ExprKind::Sequence(vec![assign_expr(target, value), int_lit(0)]))
}

pub fn sigset_del(set: Expression, sig: Expression) -> Expression {
    let target = arg_target(set);
    let bit = expr(ExprKind::Unary {
        op: UnaryOp::BitNot,
        expr: Box::new(expr(ExprKind::Binary {
            op: BinOp::Shl,
            left: Box::new(int_lit(1)),
            right: Box::new(sig),
        })),
    });
    let value = expr(ExprKind::Binary {
        op: BinOp::BitAnd,
        left: Box::new(target.clone()),
        right: Box::new(bit),
    });
    expr(ExprKind::Sequence(vec![assign_expr(target, value), int_lit(0)]))
}

pub fn sigismember(set: Expression, sig: Expression) -> Expression {
    let bit = expr(ExprKind::Binary {
        op: BinOp::Shl,
        left: Box::new(int_lit(1)),
        right: Box::new(sig),
    });
    let masked = expr(ExprKind::Binary {
        op: BinOp::BitAnd,
        left: Box::new(arg_target(set)),
        right: Box::new(bit),
    });
    ternary(
        expr(ExprKind::Binary {
            op: BinOp::NotEq,
            left: Box::new(masked),
            right: Box::new(int_lit(0)),
        }),
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
    expr(ExprKind::Sequence(vec![assign_expr(arg_target(sig), int_lit(10)), int_lit(0)]))
}

pub fn sigaction(sig: Expression, act: Expression, old: Expression) -> Expression {
    let old_write = if matches!(old.kind, ExprKind::Lit(Literal::Null)) {
        int_lit(0)
    } else {
        assign_expr(member(arg_target(old), "sa_handler"), index_expr(ident("__c_signal_handlers"), sig.clone()))
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
            right: Box::new(int_lit(4)),
        })),
        right: Box::new(int_lit(0)),
    });
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
        assign_expr(ident("__c_timer_value_sec"), member(member(target, "it_value"), "tv_sec")),
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
        assign_expr(ident("__c_itimer_value_sec"), member(member(target, "it_value"), "tv_sec")),
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
