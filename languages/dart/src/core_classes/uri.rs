//! `dart:core`'s `Uri`, as a class.
//!
//! # What this replaces: a third URL parser, written in the walker
//!
//! `Uri` was not a marker — it was a compile-time FOLD. `walker.rs` carried a
//! `DartUri` struct with its own `parse`, `split_authority`, `normalize_path_text`
//! and `default_port` in Rust, and `dart_uri_expr` baked the result into a
//! literal `ExprKind::Object` of 19 properties. That is answer substitution
//! ([[project_ruby_walker_answer_substitution]]), and it had three costs:
//!
//! 1. **It only worked for literal strings.** `Uri.parse(someVariable)` fell
//!    through to a runtime emitter whose `replace` and `normalizePath` were
//!    `pub fn …(_, _, _) {}` — empty. So the suite passed on constants while
//!    the same code on a variable did nothing.
//! 2. **`uri_decode` was `text.replace("%20", " ")`.** One escape sequence.
//!    `primitives::url::emit_percent_decode` is the real RFC 3986 codec and was
//!    already there, used by php, python, jvm and dotnet.
//! 3. **A third parser.** `primitives::url` exists precisely so `parse_url`,
//!    `urlsplit` and WHATWG are one implementation with flags.
//!
//! The class keeps the components as FIELDS and spells every derived member in
//! Dart, so `toString`, `authority`, `origin`, `normalizePath`, `replace` and
//! `resolve` lower through the shared string machinery instead of a Dart-private
//! emitter or a Rust fold.
//!
//! # Parse mode
//!
//! `ParseMode::Syntactic` with `lowercase_scheme`, chosen in
//! `emitter/core_adapter.rs`. Dart's `Uri.parse` is an RFC 3986 SPLIT: it
//! accepts a relative reference (`Uri.parse('/a/b').host == ''`), where WHATWG
//! `new URL` throws — and three tests in `uri_parsing` depend on exactly that.
//!
//! # `port` is an int with a scheme default
//!
//! Dart returns 80 for `http` and 443 for `https` when the authority carries no
//! port, and the raw component read gives `''`. The class resolves that in its
//! constructor, so `port` is the answer and not a special case at every use.

use super::builders::*;
use vybe_ast::{ClassMember, ExprKind, Expression, InterpolPart, Statement};

/// The components the constructor stores, and the profile builtin that reads
/// each one off the parsed URL. Every one of these lowers to
/// `primitives::url::emit_component_of` with the matching `UrlField`.
///
/// `scheme` is absent: it is read once and reused to resolve the default port,
/// so the constructor names it explicitly.
const COMPONENTS: &[(&str, &str)] = &[
    ("host", "__dart_url_host"),
    ("path", "__dart_url_path"),
    ("query", "__dart_url_query"),
    ("fragment", "__dart_url_fragment"),
];

/// `<fn>(<arg>)` — a profile builtin call.
fn call(name: &str, args: Vec<Expression>) -> Expression {
    Expression::with_span(
        ExprKind::Call {
            callee: Box::new(ident(name)),
            args: args
                .into_iter()
                .map(vybe_ast::Argument::positional)
                .collect(),
            optional: false,
        },
        span(),
    )
}

/// `__dart_url_decode(<expr>)` — the shared percent decoder.
fn decode(expr: Expression) -> Expression {
    call("__dart_url_decode", vec![expr])
}

/// A `String` field seeded with `''`.
fn str_field(name: &str) -> ClassMember {
    field(name, "String", str_lit(""))
}

/// `class Uri { … }`.
pub(super) fn uri() -> Statement {
    let mut members: Vec<ClassMember> = vec![
        str_field("scheme"),
        str_field("userInfo"),
        str_field("host"),
        str_field("path"),
        str_field("query"),
        str_field("fragment"),
        field("port", "int", int_lit(0)),
    ];

    // `Uri(String uri)` — ONE parse, then one component read per field.
    //
    // Percent-decoding is applied to `path`, `query`, `fragment` and `userInfo`
    // but NOT to `host`: an authority is not percent-encoded, and Dart's
    // `Uri.path` is the DECODED path (`Uri.parse('http://x/a%20b').path` is
    // `/a b`, which `uri_percent_encoded_path_segment_decoded_in_path` asserts).
    let mut body: Vec<Statement> = vec![local("__p", call("__dart_url_parse", vec![ident("uri")]))];
    body.push(set_this(
        "scheme",
        call("__dart_url_scheme", vec![ident("__p")]),
    ));
    // `userInfo` is `user[:password]`. WHATWG splits the credentials into two
    // properties and `UrlField::User`/`Pass` keep that split, so Dart's single
    // string is rejoined here — `http://user:pass@host` must read back as
    // `user:pass`, not `user`.
    body.push(local(
        "__user",
        decode(call("__dart_url_user", vec![ident("__p")])),
    ));
    body.push(local(
        "__pass",
        decode(call("__dart_url_pass", vec![ident("__p")])),
    ));
    body.push(set_this(
        "userInfo",
        ternary(
            blank(ident("__pass")),
            ident("__user"),
            interp(vec![
                InterpolPart::Expr(ident("__user")),
                InterpolPart::Text(":".to_string()),
                InterpolPart::Expr(ident("__pass")),
            ]),
        ),
    ));
    for (name, reader) in COMPONENTS {
        let read = call(reader, vec![ident("__p")]);
        let value = if *name == "host" { read } else { decode(read) };
        body.push(set_this(name, value));
    }
    // `port`: the explicit component if the authority carried one, else the
    // scheme's default. `int.parse` of `''` is not defined, so the empty case
    // is tested first.
    body.push(local("__port", call("__dart_url_port", vec![ident("__p")])));
    body.push(set_this(
        "port",
        ternary(
            blank(ident("__port")),
            default_port(),
            call_member(ident("int"), "parse", vec![ident("__port")]),
        ),
    ));
    members.push(constructor(
        vec![param("uri", Some("String"), Some(str_lit("")))],
        body,
    ));

    members.extend(derived_members());
    class("Uri", members)
}

/// `this.scheme == 'https' ? 443 : (this.scheme == 'http' ? 80 : 0)`
///
/// Dart's default-port table. Only the two schemes the language itself knows
/// have one; everything else is 0.
fn default_port() -> Expression {
    ternary(
        binary(vybe_ast::BinOp::Eq, this_field("scheme"), str_lit("https")),
        int_lit(443),
        ternary(
            binary(vybe_ast::BinOp::Eq, this_field("scheme"), str_lit("http")),
            int_lit(80),
            int_lit(0),
        ),
    )
}

/// `<a> == <b>`
fn eq(a: Expression, b: Expression) -> Expression {
    binary(vybe_ast::BinOp::Eq, a, b)
}

/// `<expr> == ''` — emptiness WITHOUT calling `.isEmpty`.
///
/// **`.isEmpty` cannot be used inside a core class body.** `StringBuffer`
/// declares `isEmpty`, the walker force-calls every `.isEmpty` read
/// (`is_dart_zero_arg_getter`), and the receiver here is a plain String the
/// compiler has no class for — so the call was diverted to StringBuffer's
/// getter, whose body reads `_vybeBuf` off a String and traps. Measured
/// 2026-08-09: `Uri` alone worked, `Uri` + `StringBuffer` in one program threw
/// `RuntimeError` out of `__Uri_ctor_0` for every URL. A comparison dispatches
/// on nothing and cannot be captured.
///
/// This is the same class-less-name hazard `defined_class_methods` creates for
/// methods ([[project_dart_duration_class_and_flat_method_set]]) — it reaches
/// getters too.
fn blank(expr: Expression) -> Expression {
    eq(expr, str_lit(""))
}

/// `!<expr>` spelled as `<expr> == false`, so it lowers through the same
/// comparison the rest of this file uses rather than a unary node.
fn not(expr: Expression) -> Expression {
    eq(expr, bool_lit(false))
}

/// Everything derived from the seven stored components.
///
/// **Every one of these is a zero-arg METHOD, and the property-shaped ones are
/// force-called by the walker.** Both halves of that are forced:
///
/// - A METHOD is required because **a Dart property getter's body cannot see
///   `this`**. Measured 2026-08-09 on a plain user class, nothing to do with
///   core classes: `String get shout => '${this.n}!'` returns `undefined!`
///   while the identical `String yell() => '${this.n}?'` returns `hi?`. Written
///   as getters, `origin` answered `undefined://undefined:undefined`. That bug
///   is why `StringBuffer`'s getters have to be routed to `dart.length` through
///   a tree `Property` leaf rather than running their own bodies — reported, not
///   worked around here.
/// - The FORCE-CALL is required because Dart writes `u.authority`, `u.origin`,
///   `u.isAbsolute`, `u.pathSegments`, `u.queryParameters` and the `has*` family
///   without parentheses. A bare read of a zero-arg method yields the function
///   object: measured `[function authority]` where `example.com:8080` was
///   wanted. `is_dart_zero_arg_getter` (`walker.rs`) is the existing mechanism
///   for exactly this, and these names join it.
///
/// `toString`, `normalizePath`, `replace` and `resolve` carry parentheses in
/// Dart source, so they need no entry there.
fn derived_members() -> Vec<ClassMember> {
    let mut out = Vec::new();

    // `authority` = `[userInfo@]host[:port]`, with a default port omitted.
    out.push(method(
        "authority",
        vec![],
        Some("String"),
        vec![ret(concat(
            concat(
                ternary(
                    blank(this_field("userInfo")),
                    str_lit(""),
                    interp(vec![
                        InterpolPart::Expr(this_field("userInfo")),
                        InterpolPart::Text("@".to_string()),
                    ]),
                ),
                this_field("host"),
            ),
            port_suffix(),
        ))],
    ));

    // `origin` = `scheme://host[:port]`. Dart throws for a schemeless URI; the
    // suite only asks for http/https, so the empty-scheme case yields ''.
    out.push(method(
        "origin",
        vec![],
        Some("String"),
        vec![ret(ternary(
            blank(this_field("scheme")),
            str_lit(""),
            concat(
                interp(vec![
                    InterpolPart::Expr(this_field("scheme")),
                    InterpolPart::Text("://".to_string()),
                    InterpolPart::Expr(this_field("host")),
                ]),
                port_suffix(),
            ),
        ))],
    ));

    // `toString` REASSEMBLES from the components — it does not echo the input.
    // That is what makes `Uri.parse(s).toString()` normalize, and it is the
    // member the suite exercises 57 times.
    out.push(method(
        "toString",
        vec![],
        Some("String"),
        vec![ret(concat(
            concat(
                concat(
                    ternary(
                        blank(this_field("scheme")),
                        str_lit(""),
                        interp(vec![
                            InterpolPart::Expr(this_field("scheme")),
                            InterpolPart::Text("://".to_string()),
                        ]),
                    ),
                    ternary(
                        blank(this_field("scheme")),
                        str_lit(""),
                        this_call("authority"),
                    ),
                ),
                this_field("path"),
            ),
            concat(
                ternary(
                    blank(this_field("query")),
                    str_lit(""),
                    interp(vec![
                        InterpolPart::Text("?".to_string()),
                        InterpolPart::Expr(this_field("query")),
                    ]),
                ),
                ternary(
                    blank(this_field("fragment")),
                    str_lit(""),
                    interp(vec![
                        InterpolPart::Text("#".to_string()),
                        InterpolPart::Expr(this_field("fragment")),
                    ]),
                ),
            ),
        ))],
    ));

    for (name, value) in [
        ("hasScheme", not(blank(this_field("scheme")))),
        ("hasAuthority", not(blank(this_field("host")))),
        ("hasQuery", not(blank(this_field("query")))),
        ("hasFragment", not(blank(this_field("fragment")))),
        ("hasEmptyPath", blank(this_field("path"))),
        ("isAbsolute", not(blank(this_field("scheme")))),
    ] {
        out.push(method(name, vec![], Some("bool"), vec![ret(value)]));
    }

    out.push(path_segments());
    out.push(query_parameters());
    out.push(normalize_path());
    out.push(replace_member());
    out.push(resolve_member());
    out.push(resolve_uri_member());
    out
}

/// `':<port>'` when the port is present and not the scheme default, else `''`.
fn port_suffix() -> Expression {
    ternary(
        eq(this_field("port"), default_port()),
        str_lit(""),
        interp(vec![
            InterpolPart::Text(":".to_string()),
            InterpolPart::Expr(this_field("port")),
        ]),
    )
}

fn interp(parts: Vec<InterpolPart>) -> Expression {
    Expression::with_span(ExprKind::Interpolation(parts), span())
}

/// `pathSegments` — the path split on `/` with empty segments dropped.
fn path_segments() -> ClassMember {
    method(
        "pathSegments",
        vec![],
        Some("List"),
        vec![
            local(
                "__parts",
                call_member(this_field("path"), "split", vec![str_lit("/")]),
            ),
            local("__out", empty_list()),
            for_in(
                "__s",
                ident("__parts"),
                vec![if_stmt(
                    not(blank(ident("__s"))),
                    vec![expr_stmt(call_member(
                        ident("__out"),
                        "add",
                        vec![ident("__s")],
                    ))],
                )],
            ),
            ret(ident("__out")),
        ],
    )
}

/// `queryParameters` — the query string as a `Map<String, String>`.
///
/// Both halves of each pair are percent-decoded through the shared codec, which
/// is what makes `?msg=hello%20world` read back as `hello world`. The walker
/// fold decoded the query with a single `replace("%20", " ")`, so anything else
/// survived escaped.
fn query_parameters() -> ClassMember {
    method(
        "queryParameters",
        vec![],
        Some("Map"),
        vec![
            local("__m", empty_map()),
            local(
                "__pairs",
                call_member(this_field("query"), "split", vec![str_lit("&")]),
            ),
            for_in(
                "__kv",
                ident("__pairs"),
                vec![if_stmt(
                    not(blank(ident("__kv"))),
                    vec![
                        local(
                            "__i",
                            call_member(ident("__kv"), "indexOf", vec![str_lit("=")]),
                        ),
                        index_set(
                            ident("__m"),
                            ternary(
                                binary(vybe_ast::BinOp::Lt, ident("__i"), int_lit(0)),
                                decode(ident("__kv")),
                                decode(call_member(
                                    ident("__kv"),
                                    "substring",
                                    vec![int_lit(0), ident("__i")],
                                )),
                            ),
                            ternary(
                                binary(vybe_ast::BinOp::Lt, ident("__i"), int_lit(0)),
                                str_lit(""),
                                decode(call_member(
                                    ident("__kv"),
                                    "substring",
                                    vec![binary(vybe_ast::BinOp::Add, ident("__i"), int_lit(1))],
                                )),
                            ),
                        ),
                    ],
                )],
            ),
            ret(ident("__m")),
        ],
    )
}

/// `normalizePath()` — RFC 3986 dot-segment removal, returning a new `Uri`.
///
/// The walker's `normalize_path_text` did this in Rust for literal URLs only.
/// Spelled here it works on any receiver.
fn normalize_path() -> ClassMember {
    method(
        "normalizePath",
        vec![],
        Some("Uri"),
        vec![
            local(
                "__parts",
                call_member(this_field("path"), "split", vec![str_lit("/")]),
            ),
            local("__keep", empty_list()),
            // `__n` shadows the list's own length on purpose. `__keep.isEmpty`
            // and `__keep.length` are both names a co-declared core class
            // claims (see `blank`), and the receiver is an untyped local — so
            // the depth is tracked in a plain int that dispatches on nothing.
            local("__n", int_lit(0)),
            for_in(
                "__s",
                ident("__parts"),
                vec![if_else(
                    eq(ident("__s"), str_lit("..")),
                    vec![if_stmt(
                        binary(vybe_ast::BinOp::Gt, ident("__n"), int_lit(0)),
                        vec![
                            expr_stmt(call_member(ident("__keep"), "removeLast", vec![])),
                            assign(
                                ident("__n"),
                                binary(vybe_ast::BinOp::Sub, ident("__n"), int_lit(1)),
                            ),
                        ],
                    )],
                    vec![if_stmt(
                        not(or(eq(ident("__s"), str_lit(".")), blank(ident("__s")))),
                        vec![
                            expr_stmt(call_member(ident("__keep"), "add", vec![ident("__s")])),
                            assign(
                                ident("__n"),
                                binary(vybe_ast::BinOp::Add, ident("__n"), int_lit(1)),
                            ),
                        ],
                    )],
                )],
            ),
            local(
                "__joined",
                concat(
                    ternary(
                        call_member(this_field("path"), "startsWith", vec![str_lit("/")]),
                        str_lit("/"),
                        str_lit(""),
                    ),
                    call_member(ident("__keep"), "join", vec![str_lit("/")]),
                ),
            ),
            ret(rebuilt(
                Some(ident("__joined")),
                None,
                None,
                None,
                None,
                None,
            )),
        ],
    )
}

/// `replace({scheme, userInfo, host, port, path, pathSegments, query, fragment})`.
///
/// Dart's signature is all-named-optional and every omitted component keeps the
/// receiver's value. Each argument therefore DEFAULTS to the field it would
/// overwrite, which makes the body a single reconstruction with no
/// per-argument branching.
///
/// `pathSegments` is the exception: it has no field to default to, and it is an
/// alternative SPELLING of `path` rather than a component of its own. It
/// defaults to the empty list and, when non-empty, wins over `path`.
fn replace_member() -> ClassMember {
    let names = [
        "scheme", "userInfo", "host", "port", "path", "query", "fragment",
    ];
    let mut params: Vec<vybe_ast::Param> = names
        .iter()
        .map(|n| param(n, None, Some(this_field(n))))
        .collect();
    params.push(param("pathSegments", None, Some(empty_list())));
    method(
        "replace",
        params,
        Some("Uri"),
        vec![
            local(
                "__path",
                ternary(
                    blank(call_member(
                        ident("pathSegments"),
                        "join",
                        vec![str_lit("/")],
                    )),
                    ident("path"),
                    concat(
                        str_lit("/"),
                        call_member(ident("pathSegments"), "join", vec![str_lit("/")]),
                    ),
                ),
            ),
            ret(rebuilt(
                Some(ident("__path")),
                Some(ident("scheme")),
                Some(ident("userInfo")),
                Some(ident("host")),
                Some(ident("port")),
                Some((ident("query"), ident("fragment"))),
            )),
        ],
    )
}

/// `resolve(ref)` / `resolveUri(ref)` — RFC 3986 reference resolution, reduced
/// to the three cases the language actually produces:
///   - an absolute reference (`a:` present) replaces everything;
///   - a rooted reference (`/…`) replaces the path;
///   - anything else joins onto the receiver's directory and normalizes.
fn resolve_member() -> ClassMember {
    resolve_named("resolve")
}

/// `resolveUri(ref)` — the same resolution, taking a `Uri` instead of a String.
///
/// The body is identical because it starts with `reference.toString()`: a `Uri`
/// reassembles to the string form its own constructor takes, so one
/// implementation serves both spellings and cannot drift from itself.
fn resolve_uri_member() -> ClassMember {
    resolve_named("resolveUri")
}

fn resolve_named(name: &str) -> ClassMember {
    method(
        name,
        vec![param("reference", None, None)],
        Some("Uri"),
        vec![
            local("__r", call_member(ident("reference"), "toString", vec![])),
            local(
                "__base",
                ternary(
                    call_member(this_field("path"), "endsWith", vec![str_lit("/")]),
                    this_field("path"),
                    call_member(
                        this_field("path"),
                        "substring",
                        vec![
                            int_lit(0),
                            binary(
                                vybe_ast::BinOp::Add,
                                call_member(this_field("path"), "lastIndexOf", vec![str_lit("/")]),
                                int_lit(1),
                            ),
                        ],
                    ),
                ),
            ),
            ret(ternary(
                call_member(ident("__r"), "contains", vec![str_lit("://")]),
                new_uri(ident("__r")),
                call_member(
                    rebuilt(
                        Some(ternary(
                            call_member(ident("__r"), "startsWith", vec![str_lit("/")]),
                            ident("__r"),
                            concat(ident("__base"), ident("__r")),
                        )),
                        None,
                        None,
                        None,
                        None,
                        None,
                    ),
                    "normalizePath",
                    vec![],
                ),
            )),
        ],
    )
}

/// `Uri(<string>)`
fn new_uri(text: Expression) -> Expression {
    Expression::with_span(
        ExprKind::New {
            class: Box::new(ident("Uri")),
            args: vec![vybe_ast::Argument::positional(text)],
        },
        span(),
    )
}

/// Rebuild a `Uri` from components, defaulting each to the receiver's own.
///
/// Reconstruction goes through the STRING form and back through the parser, so
/// a rebuilt `Uri` is parsed by the same primitive as an original one — there is
/// no second path that could disagree about what a component is.
fn rebuilt(
    path: Option<Expression>,
    scheme: Option<Expression>,
    user: Option<Expression>,
    host: Option<Expression>,
    port: Option<Expression>,
    query_fragment: Option<(Expression, Expression)>,
) -> Expression {
    let scheme = scheme.unwrap_or_else(|| this_field("scheme"));
    let user = user.unwrap_or_else(|| this_field("userInfo"));
    let host = host.unwrap_or_else(|| this_field("host"));
    let port = port.unwrap_or_else(|| this_field("port"));
    let path = path.unwrap_or_else(|| this_field("path"));
    let (query, fragment) =
        query_fragment.unwrap_or_else(|| (this_field("query"), this_field("fragment")));

    // `scheme://[user@]host[:port]path[?query][#fragment]`. The port is written
    // whenever it is non-zero and not the scheme default, which is the same
    // rule `port_suffix` applies — spelled again here because this builds from
    // the ARGUMENTS, not from the receiver's fields.
    let authority = concat(
        concat(
            ternary(
                blank(user.clone()),
                str_lit(""),
                interp(vec![
                    InterpolPart::Expr(user),
                    InterpolPart::Text("@".to_string()),
                ]),
            ),
            host,
        ),
        ternary(
            binary(
                vybe_ast::BinOp::Or,
                eq(port.clone(), int_lit(0)),
                binary(
                    vybe_ast::BinOp::Or,
                    binary(
                        vybe_ast::BinOp::And,
                        eq(scheme.clone(), str_lit("http")),
                        eq(port.clone(), int_lit(80)),
                    ),
                    binary(
                        vybe_ast::BinOp::And,
                        eq(scheme.clone(), str_lit("https")),
                        eq(port.clone(), int_lit(443)),
                    ),
                ),
            ),
            str_lit(""),
            interp(vec![
                InterpolPart::Text(":".to_string()),
                InterpolPart::Expr(port),
            ]),
        ),
    );
    let head = ternary(
        blank(scheme.clone()),
        str_lit(""),
        concat(
            interp(vec![
                InterpolPart::Expr(scheme),
                InterpolPart::Text("://".to_string()),
            ]),
            authority,
        ),
    );
    let tail = concat(
        ternary(
            blank(query.clone()),
            str_lit(""),
            interp(vec![
                InterpolPart::Text("?".to_string()),
                InterpolPart::Expr(query),
            ]),
        ),
        ternary(
            blank(fragment.clone()),
            str_lit(""),
            interp(vec![
                InterpolPart::Text("#".to_string()),
                InterpolPart::Expr(fragment),
            ]),
        ),
    );
    new_uri(concat(concat(head, path), tail))
}
