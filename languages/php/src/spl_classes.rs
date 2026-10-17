//! SPL filesystem class declarations. Methods use public adapter imports;
//! allocation, inheritance and virtual dispatch remain owned by class lowering.
use vybe_ast::{Argument, BinOp, ClassKind, ClassMember, ClassModifiers,
    ConstructorInitializerTarget, ExprKind, Expression, Literal, Modifiers,
    Param, PassBy, Statement, StmtKind, Visibility, UnaryOp, ArrayElement, CatchClause};

pub(crate) fn requires_filesystem_classes(name: &str) -> bool {
    let name = name.trim_start_matches('\\');
    // PHP class names fold ASCII case. A qualified user class must never
    // select a global builtin merely because its final component matches.
    ["SplFileInfo", "DirectoryIterator", "FilesystemIterator",
        "RecursiveDirectoryIterator", "IteratorIterator", "FilterIterator",
        "RecursiveFilterIterator", "RecursiveIteratorIterator", "NoRewindIterator",
        "RegexIterator", "RecursiveRegexIterator",
        "Traversable", "Iterator", "RecursiveIterator", "OuterIterator"]
        .iter().any(|builtin| name.eq_ignore_ascii_case(builtin))
}

fn binary(op: BinOp, left: Expression, right: Expression) -> Expression {
    Expression::new(ExprKind::Binary { op, left: Box::new(left), right: Box::new(right) })
}
fn not(value: Expression) -> Expression {
    Expression::new(ExprKind::Unary { op: UnaryOp::Not, expr: Box::new(value) })
}
fn choose(cond: Expression, then: Expression, else_: Expression) -> Expression {
    Expression::new(ExprKind::Ternary { cond: Box::new(cond), then: Box::new(then), else_: Box::new(else_) })
}
fn effect(value: Expression) -> Statement { Statement::new(StmtKind::Expr(value)) }
fn conditional(cond: Expression, then_body: Vec<Statement>) -> Statement {
    Statement::new(StmtKind::If { cond, then_body, elifs: vec![], else_body: None })
}
fn ret(value: Expression) -> Statement { Statement::new(StmtKind::Return(Some(value))) }
fn array(values: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Array(values.into_iter().map(|value| ArrayElement {
        key: None, value, spread: false, by_ref: false }).collect()))
}
fn member_call(object: Expression, name: &str, values: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call { callee: Box::new(Expression::new(ExprKind::Member {
        object: Box::new(object), field: name.into(), null_safe: false })),
        args: values.into_iter().map(Argument::positional).collect(), optional: false })
}
fn directory_init() -> Vec<Statement> {
    vec![assign("__spl_path", Expression::ident("directory")),
        assign("__spl_root", Expression::ident("directory")),
        assign("__spl_flags", Expression::ident("flags")),
        assign("__spl_entries", call("__php_spl_directory_entries", vec![Expression::ident("directory"), Expression::ident("flags")])),
        conditional(binary(BinOp::StrictEq, field("__spl_entries"), Expression::new(ExprKind::Lit(Literal::Bool(false)))), vec![
            Statement::new(StmtKind::Throw { expr: Some(Expression::new(ExprKind::New {
                class: Box::new(Expression::ident("UnexpectedValueException")),
                args: vec![Argument::positional(binary(BinOp::Concat, Expression::string("Cannot open directory: "), Expression::ident("directory")))] })), cause: None }),
        ]), assign("__spl_index", number(0))]
}

fn number(value: i64) -> Expression { Expression::new(ExprKind::Lit(Literal::Int(value))) }
fn param(name: &str, default: Option<Expression>) -> Param {
    Param { name: name.into(), type_hint: None, is_optional: default.is_some(),
        default, pass_by: PassBy::Value, is_rest: false, is_kwargs: false, is_nullable: false }
}
fn field(name: &str) -> Expression {
    Expression::new(ExprKind::Member { object: Box::new(Expression::new(ExprKind::This)),
        field: name.into(), null_safe: false })
}
fn call(name: &str, values: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call { callee: Box::new(Expression::ident(name)),
        args: values.into_iter().map(Argument::positional).collect(), optional: false })
}
fn invoke(name: &str, values: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call { callee: Box::new(field(name)),
        args: values.into_iter().map(Argument::positional).collect(), optional: false })
}
fn assign(name: &str, value: Expression) -> Statement {
    Statement::new(StmtKind::Expr(Expression::new(ExprKind::Assign {
        target: Box::new(field(name)), value: Box::new(value) })))
}
fn member_field(name: &str, value: Expression) -> ClassMember {
    ClassMember::Field { name: name.into(), type_hint: None, init: Some(value),
        modifiers: Modifiers { visibility: Visibility::Protected, ..Modifiers::default() },
        with_events: false, array_bounds: None, storage: None }
}
fn method(name: &str, params: Vec<Param>, body: Vec<Statement>) -> ClassMember {
    // These are adapter contracts, not guesses based on user method names.
    // Preserve mixed/false-capable results instead of claiming a narrower type.
    let return_type = match name {
        "valid" | "isDot" | "isDir" | "isFile" | "isLink" | "isReadable"
        | "isWritable" | "hasChildren" | "__spl_query_children" => Some("bool".into()),
        "getPathname" | "getFilename" | "getPath" | "getExtension"
        | "getSubPath" | "getSubPathname" | "__toString" => Some("string".into()),
        _ => None,
    };
    ClassMember::Method(Box::new(Statement::new(StmtKind::FunctionDecl {
        name: name.into(), params, return_type, body, modifiers: Modifiers::default(),
        handles: vec![], is_async: false, is_generator: false, is_sub: false })))
}
fn interface_method(name: &str, return_type: &str) -> ClassMember {
    ClassMember::Method(Box::new(Statement::new(StmtKind::FunctionDecl {
        name: name.into(), params: vec![], return_type: Some(return_type.into()), body: vec![],
        modifiers: Modifiers { is_abstract: true, ..Modifiers::default() },
        handles: vec![], is_async: false, is_generator: false, is_sub: false })))
}
fn returning(name: &str, value: Expression) -> ClassMember {
    method(name, vec![], vec![Statement::new(StmtKind::Return(Some(value)))])
}
fn constructor(params: Vec<Param>, body: Vec<Statement>) -> ClassMember {
    ClassMember::Constructor { name: None, params, body, base_args: None,
        initializer_target: ConstructorInitializerTarget::Base, visibility: Visibility::Public }
}
fn class(name: &str, parents: &[&str], interfaces: &[&str], members: Vec<ClassMember>, kind: ClassKind) -> Statement {
    Statement::new(StmtKind::ClassDecl { name: name.into(),
        parents: parents.iter().map(|s| (*s).into()).collect(),
        interfaces: interfaces.iter().map(|s| (*s).into()).collect(), members,
        modifiers: ClassModifiers { kind,
            linkage: vybe_ast::ClassLinkage::Library { identity: format!("php:spl:v1:{name}") },
            ..ClassModifiers::default() }, decorators: vec![] })
}
fn new_file(path: Expression) -> Expression {
    Expression::new(ExprKind::New { class: Box::new(Expression::ident("SplFileInfo")),
        args: vec![Argument::positional(path)] })
}

pub(crate) fn declarations() -> Vec<Statement> {
    let mut result = vec![
        class("Traversable", &[], &[], vec![], ClassKind::Interface),
        class("Iterator", &["Traversable"], &[], vec![
            interface_method("current", "mixed"), interface_method("key", "mixed"),
            interface_method("next", "void"), interface_method("rewind", "void"),
            interface_method("valid", "bool"),
        ], ClassKind::Interface),
        class("RecursiveIterator", &["Iterator"], &[], vec![
            interface_method("hasChildren", "bool"),
            interface_method("getChildren", "RecursiveIterator|null"),
        ], ClassKind::Interface),
    ];
    let mut info = vec![member_field("__spl_path", Expression::string("")),
        constructor(vec![param("filename", None)], vec![assign("__spl_path", Expression::ident("filename"))]),
        returning("getPathname", field("__spl_path")),
        returning("getFilename", call("basename", vec![invoke("getPathname", vec![])])),
        returning("getPath", call("dirname", vec![invoke("getPathname", vec![])])),
        returning("getRealPath", call("realpath", vec![invoke("getPathname", vec![])])),
        returning("__toString", invoke("getPathname", vec![])),
        returning("getExtension", call("__php_spl_extension", vec![invoke("getPathname", vec![])])),
    ];
    for (method_name, builtin) in [("isDir", "is_dir"), ("isFile", "is_file"),
        ("isLink", "is_link"), ("isReadable", "is_readable"),
        ("isWritable", "is_writable"), ("getSize", "filesize"),
        ("getMTime", "filemtime"), ("getATime", "fileatime"), ("getCTime", "filectime")] {
        info.push(returning(method_name, call(builtin, vec![invoke("getPathname", vec![])])));
    }
    result.push(class("SplFileInfo", &[], &[], info, ClassKind::Class));
    let pathname = || call("__php_spl_directory_entry", vec![field("__spl_entries"), field("__spl_index")]);
    let directory = vec![
        member_field("__spl_entries", Expression::new(ExprKind::Array(vec![]))),
        member_field("__spl_index", number(0)), member_field("__spl_flags", number(0)),
        member_field("__spl_root", Expression::string("")),
        constructor(vec![param("directory", None), param("flags", Some(number(0)))], directory_init()),
        returning("getPathname", pathname()), returning("current", Expression::new(ExprKind::This)),
        returning("key", field("__spl_index")),
        returning("valid", call("__php_spl_directory_valid", vec![field("__spl_entries"), field("__spl_index")])),
        returning("isDot", call("__php_spl_is_dot", vec![pathname()])),
        method("rewind", vec![], vec![assign("__spl_index", number(0))]),
        method("next", vec![], vec![assign("__spl_index", Expression::new(ExprKind::Binary {
            op: BinOp::Add, left: Box::new(field("__spl_index")), right: Box::new(number(1)) }))]),
        method("seek", vec![param("position", None)], vec![assign("__spl_index", Expression::ident("position"))]),
    ];
    result.push(class("DirectoryIterator", &["SplFileInfo"], &["Iterator"], directory, ClassKind::Class));
    let mut filesystem = vec![
        returning("current", choose(binary(BinOp::Eq, binary(BinOp::BitAnd, field("__spl_flags"), number(240)), number(32)), pathname(),
            choose(binary(BinOp::Eq, binary(BinOp::BitAnd, field("__spl_flags"), number(240)), number(16)), Expression::new(ExprKind::This), new_file(pathname())))),
        returning("key", choose(binary(BinOp::NotEq, binary(BinOp::BitAnd, field("__spl_flags"), number(256)), number(0)),
            invoke("getFilename", vec![]), pathname())),
        returning("getFlags", field("__spl_flags")),
        method("setFlags", vec![param("flags", None)], vec![assign("__spl_flags", Expression::ident("flags"))]),
        constructor(vec![param("directory", None), param("flags", Some(number(4096)))], directory_init()),
    ];
    for (name, value) in [("CURRENT_AS_FILEINFO", 0), ("CURRENT_AS_SELF", 16), ("CURRENT_AS_PATHNAME", 32),
        ("CURRENT_MODE_MASK", 240), ("KEY_AS_PATHNAME", 0), ("KEY_AS_FILENAME", 256),
        ("KEY_MODE_MASK", 3840), ("FOLLOW_SYMLINKS", 16384), ("SKIP_DOTS", 4096), ("UNIX_PATHS", 8192)] {
        filesystem.push(ClassMember::Const { name: name.into(), type_hint: None,
            value: number(value), visibility: Visibility::Public });
    }
    result.push(class("FilesystemIterator", &["DirectoryIterator"], &[], filesystem, ClassKind::Class));
    let recursive = vec![
        constructor(vec![param("directory", None), param("flags", Some(number(0)))], directory_init()),
        method("hasChildren", vec![param("allowLinks", Some(Expression::new(ExprKind::Lit(Literal::Bool(false)))))],
            vec![Statement::new(StmtKind::Return(Some(call("__php_spl_has_children", vec![pathname(), field("__spl_flags"), Expression::ident("allowLinks")]))))]),
        method("getChildren", vec![], vec![
            effect(Expression::new(ExprKind::Assign { target: Box::new(Expression::ident("child")),
                value: Box::new(Expression::new(ExprKind::New { class: Box::new(call("__php_get_class", vec![Expression::new(ExprKind::This)])),
                    args: vec![Argument::positional(pathname()), Argument::positional(field("__spl_flags"))] })) })),
            effect(Expression::new(ExprKind::Assign { target: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident("child")), field: "__spl_root".into(), null_safe: false })), value: Box::new(field("__spl_root")) })),
            ret(Expression::ident("child")),
        ]),
        returning("getSubPath", call("__php_spl_relative_path", vec![field("__spl_root"), field("__spl_path"), field("__spl_flags")])),
        returning("getSubPathname", call("__php_spl_relative_path", vec![field("__spl_root"), pathname(), field("__spl_flags")])),
    ];
    result.push(class("RecursiveDirectoryIterator", &["FilesystemIterator"], &["RecursiveIterator"], recursive, ClassKind::Class));
    result.extend(iterator_declarations());
    result
}

fn iterator_declarations() -> Vec<Statement> {
    let inner = || field("__spl_inner");
    let delegate = |name| member_call(inner(), name, vec![]);
    let mut iterator = vec![member_field("__spl_inner", Expression::null()),
        constructor(vec![param("iterator", None)], vec![assign("__spl_inner", Expression::ident("iterator"))]),
        returning("getInnerIterator", inner())];
    for name in ["rewind", "next", "key", "current", "valid"] {
        iterator.push(returning(name, delegate(name)));
    }
    let skip = || invoke("__spl_skip_rejected", vec![]);
    let filter = vec![
        method("__spl_skip_rejected", vec![], vec![Statement::new(StmtKind::While {
            cond: binary(BinOp::And, delegate("valid"), not(invoke("accept", vec![]))),
            body: vec![effect(delegate("next"))], else_body: None })]),
        method("rewind", vec![], vec![effect(delegate("rewind")), effect(skip())]),
        method("next", vec![], vec![effect(delegate("next")), effect(skip())]),
    ];
    let recursive_filter = vec![returning("hasChildren", delegate("hasChildren")),
        returning("getChildren", Expression::new(ExprKind::New {
            class: Box::new(call("__php_get_class", vec![Expression::new(ExprKind::This)])),
            args: vec![Argument::positional(delegate("getChildren"))] }))];
    vec![class("OuterIterator", &["Iterator"], &[], vec![
            interface_method("getInnerIterator", "Iterator|null"),
        ], ClassKind::Interface),
        class("IteratorIterator", &[], &["OuterIterator"], iterator, ClassKind::Class),
        class("FilterIterator", &["IteratorIterator"], &[], filter, ClassKind::Class),
        class("RecursiveFilterIterator", &["FilterIterator"], &["RecursiveIterator"], recursive_filter, ClassKind::Class),
        class("NoRewindIterator", &["IteratorIterator"], &[], vec![method("rewind", vec![], vec![])], ClassKind::Class),
        regex_iterator_declaration(), recursive_regex_iterator_declaration(),
        recursive_iterator_declaration()]
}

fn regex_mode_check(value: Expression) -> Statement {
    let valid = (1..=4).fold(binary(BinOp::StrictEq, value.clone(), number(0)), |valid, mode| {
        binary(BinOp::Or, valid, binary(BinOp::StrictEq, value.clone(), number(mode)))
    });
    conditional(not(valid), vec![Statement::new(StmtKind::Throw {
        expr: Some(Expression::new(ExprKind::New {
            class: Box::new(Expression::ident("InvalidArgumentException")),
            args: vec![Argument::positional(Expression::string("Invalid RegexIterator mode"))],
        })), cause: None,
    })])
}

fn regex_iterator_declaration() -> Statement {
    let inner = || field("__spl_inner");
    let delegate = |name| member_call(inner(), name, vec![]);
    let use_key = || binary(BinOp::BitAnd, field("__spl_regex_flags"), number(1));
    let entry = |index| Expression::new(ExprKind::Index {
        object: Box::new(field("__spl_regex_entry")), index: Box::new(number(index)), null_safe: false,
    });
    let mut members = vec![
        member_field("__spl_regex", Expression::string("")),
        member_field("__spl_regex_mode", number(0)),
        member_field("__spl_regex_flags", number(0)),
        member_field("__spl_regex_preg_flags", number(0)),
        member_field("__spl_regex_entry", array(vec![])),
        member_field("__spl_regex_key", Expression::null()),
        member_field("__spl_regex_value", Expression::null()),
        ClassMember::Field { name: "replacement".into(), type_hint: Some("string".into()),
            init: Some(Expression::null()), modifiers: Modifiers::default(),
            with_events: false, array_bounds: None, storage: None },
        constructor(vec![param("iterator", None), param("pattern", None),
            param("mode", Some(number(0))), param("flags", Some(number(0))),
            param("pregFlags", Some(number(0)))], vec![
            regex_mode_check(Expression::ident("mode")),
            conditional(not(call("__php_regex_valid_pattern", vec![Expression::ident("pattern")])),
                vec![Statement::new(StmtKind::Throw {
                    expr: Some(Expression::new(ExprKind::New {
                        class: Box::new(Expression::ident("InvalidArgumentException")),
                        args: vec![Argument::positional(Expression::string("Invalid regular expression"))],
                    })), cause: None,
                })]),
            assign("__spl_inner", Expression::ident("iterator")),
            assign("__spl_regex", Expression::ident("pattern")),
            assign("__spl_regex_mode", Expression::ident("mode")),
            assign("__spl_regex_flags", Expression::ident("flags")),
            assign("__spl_regex_preg_flags", Expression::ident("pregFlags")),
        ]),
        method("__spl_regex_accept", vec![], vec![
            conditional(not(delegate("valid")), vec![ret(Expression::new(ExprKind::Lit(Literal::Bool(false))))]),
            assign("__spl_regex_key", delegate("key")),
            assign("__spl_regex_value", delegate("current")),
            conditional(binary(BinOp::And, not(use_key()), call("is_array", vec![field("__spl_regex_value")])),
                vec![ret(Expression::new(ExprKind::Lit(Literal::Bool(false))))]),
            assign("__spl_regex_entry", call("__php_regex_iterator_entry", vec![
                field("__spl_regex"), call("strval", vec![choose(use_key(),
                    field("__spl_regex_key"), field("__spl_regex_value"))]),
                field("__spl_regex_mode"), field("__spl_regex_flags"),
                field("__spl_regex_preg_flags"), field("replacement"),
            ])),
            conditional(binary(BinOp::StrictEq, field("__spl_regex_mode"), number(4)), vec![
                Statement::new(StmtKind::If { cond: use_key(),
                    then_body: vec![assign("__spl_regex_key", entry(1))], elifs: vec![],
                    else_body: Some(vec![assign("__spl_regex_value", entry(1))]) }),
            ]),
            conditional(binary(BinOp::And,
                not(binary(BinOp::StrictEq, field("__spl_regex_mode"), number(0))),
                not(binary(BinOp::StrictEq, field("__spl_regex_mode"), number(4)))),
                vec![assign("__spl_regex_value", entry(1))]),
            ret(entry(0)),
        ]),
        returning("accept", invoke("__spl_regex_accept", vec![])),
        returning("current", field("__spl_regex_value")),
        returning("key", field("__spl_regex_key")),
        returning("getRegex", field("__spl_regex")),
        returning("getMode", field("__spl_regex_mode")),
        returning("getFlags", field("__spl_regex_flags")),
        returning("getPregFlags", field("__spl_regex_preg_flags")),
        method("setMode", vec![param("mode", None)], vec![
            regex_mode_check(Expression::ident("mode")), assign("__spl_regex_mode", Expression::ident("mode"))]),
        method("setFlags", vec![param("flags", None)], vec![assign("__spl_regex_flags", Expression::ident("flags"))]),
        method("setPregFlags", vec![param("pregFlags", None)], vec![assign("__spl_regex_preg_flags", Expression::ident("pregFlags"))]),
    ];
    // Keys are cached independently of transformed values, including USE_KEY
    // replacement. FilterIterator owns lazy rewind/next and rejection skipping.
    for (name, value) in [("MATCH", 0), ("GET_MATCH", 1), ("ALL_MATCHES", 2),
        ("SPLIT", 3), ("REPLACE", 4), ("USE_KEY", 1), ("INVERT_MATCH", 2)] {
        members.push(ClassMember::Const { name: name.into(), type_hint: None,
            value: number(value), visibility: Visibility::Public });
    }
    class("RegexIterator", &["FilterIterator"], &[], members, ClassKind::Class)
}

fn recursive_regex_iterator_declaration() -> Statement {
    let delegate = |name| member_call(field("__spl_inner"), name, vec![]);
    class("RecursiveRegexIterator", &["RegexIterator"], &["RecursiveIterator"], vec![
        method("accept", vec![], vec![
            conditional(delegate("hasChildren"), vec![
                assign("__spl_regex_key", delegate("key")),
                assign("__spl_regex_value", delegate("current")),
                ret(Expression::new(ExprKind::Lit(Literal::Bool(true)))),
            ]),
            ret(invoke("__spl_regex_accept", vec![])),
        ]),
        returning("hasChildren", delegate("hasChildren")),
        returning("getChildren", Expression::new(ExprKind::New {
            class: Box::new(Expression::ident("RecursiveRegexIterator")),
            args: vec![delegate("getChildren"), field("__spl_regex"), field("__spl_regex_mode"),
                field("__spl_regex_flags"), field("__spl_regex_preg_flags")]
                .into_iter().map(Argument::positional).collect(),
        })),
    ], ClassKind::Class)
}

// Each frame advances through before-entry, descent, after-child and next-entry
// phases. No eager flattening: user filters/overrides execute only as traversed.
fn recursive_iterator_declaration() -> Statement {
    let stack = || field("__spl_stack");
    let frames = || call("count", vec![stack()]);
    let top = || call("end", vec![stack()]);
    let top_call = |name| member_call(top(), name, vec![]);
    let phase = || call("end", vec![field("__spl_phases")]);
    let depth = || choose(binary(BinOp::Gt, frames(), number(0)), binary(BinOp::Sub, frames(), number(1)), number(0));
    let children_allowed = || binary(BinOp::Or, binary(BinOp::Lt, field("__spl_max_depth"), number(0)),
        binary(BinOp::Lt, depth(), field("__spl_max_depth")));
    let has_children = || binary(BinOp::And, children_allowed(), field("__spl_child_available"));
    let set_phase = |value| effect(call("__php_spl_set_top", vec![field("__spl_phases"), number(value)]));
    let push_child = || vec![set_phase(2), effect(invoke("__spl_descend", vec![]))];
    let yielded = || ret(Expression::null());
    let catch_disabled = || binary(BinOp::Eq, binary(BinOp::BitAnd, field("__spl_flags"), number(16)), number(0));
    let rethrow = || conditional(catch_disabled(), vec![Statement::new(StmtKind::Throw {
        expr: Some(Expression::ident("child_error")), cause: None })]);
    let query_children = method("__spl_query_children", vec![], vec![Statement::new(StmtKind::Try {
        body: vec![ret(top_call("hasChildren"))],
        catches: vec![CatchClause { types: vec!["Throwable".into()], var_name: Some("child_error".into()),
            stack_var: None, when_clause: None,
            body: vec![rethrow(), ret(Expression::new(ExprKind::Lit(Literal::Bool(false))))] }],
        else_body: None, finally: None,
    })]);
    let descend = method("__spl_descend", vec![], vec![
        Statement::new(StmtKind::Try {
            body: vec![effect(Expression::new(ExprKind::Assign {
                target: Box::new(Expression::ident("child")), value: Box::new(top_call("getChildren")) }))],
            catches: vec![CatchClause { types: vec!["Throwable".into()], var_name: Some("child_error".into()),
                stack_var: None, when_clause: None,
                body: vec![rethrow(), set_phase(3), yielded()] }], else_body: None, finally: None,
        }),
        effect(call("array_push", vec![stack(), Expression::ident("child")])),
        effect(call("array_push", vec![field("__spl_phases"), number(0)])),
        // CATCH_GET_CHILD does not suppress a child's rewind failure.
        effect(top_call("rewind")),
    ]);
    let select = method("__spl_select", vec![], vec![Statement::new(StmtKind::While {
        cond: binary(BinOp::Gt, frames(), number(0)), else_body: None,
        body: vec![Statement::new(StmtKind::If { cond: not(top_call("valid")),
            then_body: vec![effect(call("array_pop", vec![stack()])), effect(call("array_pop", vec![field("__spl_phases")]))],
            elifs: vec![(binary(BinOp::Eq, phase(), number(0)), vec![set_phase(1),
                assign("__spl_child_available", invoke("__spl_query_children", vec![])),
                conditional(binary(BinOp::Eq, field("__spl_mode"), number(1)), vec![yielded()]),
                // A depth cutoff prevents descent; it does not turn a
                // directory into a leaf for LEAVES_ONLY traversal.
                conditional(not(field("__spl_child_available")), vec![set_phase(3), yielded()])]),
                (binary(BinOp::Eq, phase(), number(1)), vec![Statement::new(StmtKind::If {
                    cond: has_children(), then_body: push_child(), elifs: vec![],
                    else_body: Some(vec![set_phase(3),
                        conditional(binary(BinOp::Eq, field("__spl_mode"), number(2)), vec![yielded()])]) })]),
                (binary(BinOp::Eq, phase(), number(2)), vec![set_phase(3),
                    conditional(binary(BinOp::Eq, field("__spl_mode"), number(2)), vec![yielded()])])],
            else_body: Some(vec![effect(top_call("next")), set_phase(0)]) })],
    })]);
    let mut members = vec![member_field("__spl_stack", array(vec![])), member_field("__spl_phases", array(vec![])),
        member_field("__spl_max_depth", number(-1)),
        member_field("__spl_flags", number(0)),
        member_field("__spl_child_available", Expression::new(ExprKind::Lit(Literal::Bool(false)))),
        member_field("__spl_mode", number(0)), member_field("__spl_inner", Expression::null()),
        constructor(vec![param("iterator", None), param("mode", Some(number(0))), param("flags", Some(number(0)))],
            vec![assign("__spl_inner", Expression::ident("iterator")), assign("__spl_mode", Expression::ident("mode")),
                assign("__spl_flags", Expression::ident("flags"))]),
        method("rewind", vec![], vec![assign("__spl_stack", array(vec![field("__spl_inner")])),
            assign("__spl_phases", array(vec![number(0)])), effect(top_call("rewind")), effect(invoke("__spl_select", vec![]))]),
        method("next", vec![], vec![effect(invoke("__spl_select", vec![]))]),
        returning("valid", binary(BinOp::Gt, frames(), number(0))),
        returning("current", choose(binary(BinOp::Gt, frames(), number(0)), top_call("current"), Expression::null())),
        returning("key", choose(binary(BinOp::Gt, frames(), number(0)), top_call("key"), Expression::null())),
        returning("getDepth", depth()), returning("getMaxDepth", field("__spl_max_depth")),
        method("setMaxDepth", vec![param("maxDepth", Some(number(-1)))], vec![
            conditional(binary(BinOp::Lt, Expression::ident("maxDepth"), number(-1)), vec![
                Statement::new(StmtKind::Throw { expr: Some(Expression::new(ExprKind::New {
                    class: Box::new(Expression::ident("ValueError")), args: vec![Argument::positional(Expression::string("Maximum depth must be greater than or equal to -1"))] })), cause: None })]),
            assign("__spl_max_depth", Expression::ident("maxDepth"))]),
        returning("getInnerIterator", choose(binary(BinOp::Gt, frames(), number(0)), top(), field("__spl_inner"))),
        query_children, descend, select];
    for (name, value) in [("LEAVES_ONLY", 0), ("SELF_FIRST", 1), ("CHILD_FIRST", 2), ("CATCH_GET_CHILD", 16)] {
        members.push(ClassMember::Const { name: name.into(), type_hint: None, value: number(value), visibility: Visibility::Public });
    }
    class("RecursiveIteratorIterator", &[], &["OuterIterator"], members, ClassKind::Class)
}
