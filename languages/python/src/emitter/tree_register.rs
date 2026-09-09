//! Python namespace-tree registration.
//!
//! Mirrors `languages/php/src/tree_register.rs`: the LANGUAGE contributes DATA
//! — its own profile tables, the same ones its emit dispatch executes — to the
//! shared tree in `vybe_compiler::primitives::namespaces`. Resolution logic lives only in
//! the common resolver; nothing in the VM or the common compiler changes.
//!
//! The one difference from PHP is what a DOTTED profile key means. PHP skips
//! them (`$obj->method` is receiver dispatch), but in Python `"os.stat"`,
//! `"os.path.join"` and `"math.pi"` ARE module members, so each dotted key
//! nests into real subtrees — `os` → `path` → `join`. That is what makes
//! `from os import stat` / `from math import pi` resolve generically, instead
//! of needing a hand-written `[[esm_default]] kind = "module-export"` row per
//! name (the json rows are exactly that workaround).
//!
//! Leaf kinds come from the profile entry:
//! - `emit = "common:python.<fn>"` → `CommonEmit`
//! - `emit = "host:<module>:<fn>"` → `Fn`
//! - `[namespace_constants]` → `Const`
//! - opcode/intrinsic/print builtins have no process-global target — skipped.
//!
//! Python is CASE-SENSITIVE, so keys keep their exact source casing
//! (`OrderedDict`, `NamedTemporaryFile`).
//!
//! ⛔ That USED to require bypassing the `namespace()` helper, which asserted
//! keys were lowercase-canonical. **This registrar was right and the invariant
//! was wrong**: the rule was a shortcut from the era when several namespaces
//! and resolvers were being unified, and it served the five case-insensitive
//! languages at the expense of the twelve others. The assertion is gone, tree
//! lookups match EXACT first and fold only on a miss, and every registrar now
//! keeps the case the source wrote — see
//! `documentation/casesensitivityplan.md`. Building the `Subtree` directly
//! here is now an ordinary choice rather than a workaround.

use std::collections::BTreeMap;
use std::sync::Once;

use vybe_compiler::primitives::namespaces::{self, CtorSpec, NamespaceNode, Subtree};
use vybe_runtime::Value;
use vybe_runtime::profile::{BuiltinEmit, ConstantValue, parse_profile};

/// Insert `leaf` at a dotted path (`["path", "join"]` under root `os`),
/// creating intermediate namespaces. An existing entry wins — first
/// registration is authoritative, as in the other registrars.
fn insert_path(root: &mut Subtree, path: &[&str], leaf: NamespaceNode) {
    let Some((last, parents)) = path.split_last() else {
        return;
    };
    let mut current = root;
    for seg in parents {
        let entry = current
            .entry((*seg).to_string())
            .or_insert_with(|| NamespaceNode::Namespace(Subtree::new()));
        match entry {
            NamespaceNode::Namespace(children) => current = children,
            // A leaf already occupies this segment — it cannot also be a
            // namespace, so leave the working entry alone.
            _ => return,
        }
    }
    current.entry((*last).to_string()).or_insert(leaf);
}

/// Register the Python surface. Idempotent; first call wins.
pub fn register_namespace_tree() {
    static ONCE: Once = Once::new();
    ONCE.call_once(|| {
        register_from_profile();
    });
}

/// Everything derivable from the profile — no hand-maintained module list.
fn register_from_profile() {
    let Ok(profile) = parse_profile(crate::profile_source()) else {
        return;
    };

    let mut roots: BTreeMap<String, Subtree> = BTreeMap::new();
    // Hand-written surfaces are MODULES like any other, so they join the same
    // map rather than claiming roots of their own.
    roots.insert("collections".to_string(), collections_subtree());
    roots.insert("calendar".to_string(), calendar_subtree());
    roots.insert("dataclasses".to_string(), dataclasses_subtree());
    roots.insert("doctest".to_string(), doctest_subtree());
    roots.insert("dis".to_string(), dis_subtree());
    roots.insert("email".to_string(), email_subtree());
    roots.insert("pydoc".to_string(), pydoc_subtree());
    roots.insert("symtable".to_string(), symtable_subtree());
    roots.insert("token".to_string(), token_subtree(false));
    roots.insert("tokenize".to_string(), token_subtree(true));
    for (module, tree) in core_class_subtrees() {
        roots.entry(module).or_default().extend(tree);
    }
    let mut add = |key: &str, leaf: NamespaceNode| {
        let segments: Vec<&str> = key.split('.').collect();
        // A bare builtin (`len`, `print`) is not module surface.
        if segments.len() < 2 {
            return;
        }
        let root = roots.entry(segments[0].to_string()).or_default();
        insert_path(root, &segments[1..], leaf);
    };

    for (name, def) in &profile.builtins {
        match &def.emit {
            BuiltinEmit::Common(op) => add(name, NamespaceNode::CommonEmit(op.clone())),
            BuiltinEmit::HostCall(module, func) => add(name, namespaces::host_fn(module, func)),
            _ => {}
        }
    }

    // `[namespace_constants]` are VALUES (`math.pi`, `math.inf`) — precisely
    // the members a named import binds. Without them `from math import pi` has
    // nothing to resolve to and lands as `nan`.
    for (name, value) in &profile.namespace_constants {
        let node = match value {
            ConstantValue::Bool(b) => NamespaceNode::Const(Value::Bool(*b)),
            ConstantValue::Float(f) => NamespaceNode::Const(Value::F64(*f)),
            // `inf`/`nan` are spelled as strings in the profile but ARE floats.
            ConstantValue::Str(s) => match s.as_str() {
                "Infinity" => NamespaceNode::Const(Value::F64(f64::INFINITY)),
                "-Infinity" => NamespaceNode::Const(Value::F64(f64::NEG_INFINITY)),
                "NaN" => NamespaceNode::Const(Value::F64(f64::NAN)),
                _ => NamespaceNode::Const(Value::String(std::sync::Arc::from(s.as_str()))),
            },
        };
        add(name, node);
    }

    // ECMA-shared math names. The host exports them under `ecma:math`, which
    // `mount_host_exports` mounts LOWERCASED as `ecma.math.<name>` — so an
    // `Alias` leaf reaches them with no per-name profile row. namespaceplan.md
    // §"Source-name ≠ canonical-name": `python.json.dumps =
    // Alias(ecma.json.stringify)` is exactly this shape. `erf`/`gamma` family
    // have real `common:math.*64` profile leaves, so they are derived above.
    for name in ["hypot"] {
        // `Path` is a dotted string.
        let target = format!("ecma.math.{name}");
        let root = roots.entry("math".to_string()).or_default();
        insert_path(root, &[name], NamespaceNode::Alias(target));
    }

    // Python spells one leaf under two module names. namespaceplan.md
    // §"Source-name ≠ canonical-name": an `Alias` NAMES the existing typed
    // leaf instead of restating it, so the tree holds exactly one CRC-32.
    for (module, member, target) in [("zlib", "crc32", "python.binascii.crc32")] {
        let root = roots.entry(module.to_string()).or_default();
        insert_path(root, &[member], NamespaceNode::Alias(target.to_string()));
    }

    // ONE root, the way `platforms/jvm` owns `jvm.java.*`, kotlin owns
    // `kotlin.*` and dotnet owns `dotnet.*` (namespaceplan.md §"Profile mounts").
    //
    // This used to ALSO register each module as a bare global root — "same
    // data, two mount points" — which put `str`, `float`, `array`, `math`,
    // `os`, `time` and 25 more into the one global tree that every language
    // shares. The source-level names are not lost: the profile mounts them
    // with `[[esm_default]] kind = "tree-mount"`, so `os.path.join` resolves
    // through the COMMON resolver against `python.os.path.join` instead of a
    // root only python could have registered.
    let mut python_root = Subtree::new();
    for (root, tree) in roots {
        python_root.insert(root, NamespaceNode::Namespace(tree));
    }
    namespaces::register_namespace_tree("python", NamespaceNode::Namespace(python_root));
}

/// `collections` constructors. Hand-written because these are not profile
/// builtins: `deque`/`OrderedDict` ARE plain ecma constructors, and
/// `Counter`/`defaultdict` need custom construction.
///
/// Instance methods (`rotate`, `move_to_end`, `most_common`, …) are NEVER
/// resolved through the tree — member dispatch is receiver-based.
fn collections_subtree() -> Subtree {
    let mut root = Subtree::new();
    root.insert(
        "deque".to_string(),
        namespaces::host_fn("ecma:array", "from"),
    );
    root.insert(
        "OrderedDict".to_string(),
        NamespaceNode::CommonEmit("python.ordereddict_new".to_string()),
    );
    root.insert(
        "Counter".to_string(),
        NamespaceNode::CommonEmit("python.counter_new".to_string()),
    );
    root.insert(
        "defaultdict".to_string(),
        NamespaceNode::CommonEmit("python.defaultdict_new".to_string()),
    );
    root
}

fn dataclasses_subtree() -> Subtree {
    let mut root = Subtree::new();
    for (name, emit) in [
        ("is_dataclass", "python.is_dataclass"),
        ("asdict", "python.dataclass_asdict"),
        ("astuple", "python.dataclass_astuple"),
        ("fields", "python.dataclass_fields"),
    ] {
        root.insert(name.to_string(), NamespaceNode::CommonEmit(emit.to_string()));
    }
    root.insert(
        "MISSING".to_string(),
        NamespaceNode::Const(Value::String(std::sync::Arc::from("MISSING"))),
    );
    root.insert(
        "KW_ONLY".to_string(),
        NamespaceNode::Const(Value::String(std::sync::Arc::from("KW_ONLY"))),
    );
    root.insert(
        "InitVar".to_string(),
        NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics: Subtree::new(),
            methods: Subtree::new(),
            member_returns: BTreeMap::new(),
        },
    );
    root.insert(
        "FrozenInstanceError".to_string(),
        NamespaceNode::CommonEmit("python.exc.FrozenInstanceError".to_string()),
    );
    root
}

fn calendar_type(name: &str, ctor: &str, formatmonth: Option<&str>) -> NamespaceNode {
    let mut methods = BTreeMap::from([
        (
            "itermonthdays".to_string(),
            NamespaceNode::CommonEmit("python.calendar_itermonthdays".to_string()),
        ),
        (
            "itermonthdates".to_string(),
            NamespaceNode::CommonEmit("python.calendar_itermonthdays".to_string()),
        ),
        (
            "itermonthdays2".to_string(),
            NamespaceNode::CommonEmit("python.calendar_itermonthdays2".to_string()),
        ),
        (
            "monthdayscalendar".to_string(),
            NamespaceNode::CommonEmit("python.calendar_monthcalendar".to_string()),
        ),
        (
            "yeardayscalendar".to_string(),
            NamespaceNode::CommonEmit("python.calendar_yeardayscalendar".to_string()),
        ),
    ]);
    if let Some(emit) = formatmonth {
        methods.insert(
            "formatmonth".to_string(),
            NamespaceNode::CommonEmit(emit.to_string()),
        );
        methods.insert(
            "prmonth".to_string(),
            NamespaceNode::CommonEmit(emit.to_string()),
        );
    }
    NamespaceNode::Type {
        ctor: Some(CtorSpec {
            params: vec!["firstweekday".to_string()],
            fields: vec!["firstweekday".to_string()],
            ancestry: vec![name.to_string()],
            ..Default::default()
        }),
        ctor_call: Some(Box::new(NamespaceNode::CommonEmit(ctor.to_string()))),
        statics: Subtree::new(),
        methods,
        member_returns: BTreeMap::new(),
    }
}

fn calendar_subtree() -> Subtree {
    let mut root = Subtree::new();
    root.insert(
        "Calendar".to_string(),
        calendar_type("Calendar", "python.calendar_new", None),
    );
    root.insert(
        "TextCalendar".to_string(),
        calendar_type(
            "TextCalendar",
            "python.calendar_text_new",
            Some("python.calendar_text_formatmonth"),
        ),
    );
    root.insert(
        "HTMLCalendar".to_string(),
        calendar_type(
            "HTMLCalendar",
            "python.calendar_html_new",
            Some("python.calendar_html_formatmonth"),
        ),
    );
    root
}

fn simple_python_type(name: &str, ctor_emit: &str, methods: &[(&str, &str)]) -> NamespaceNode {
    NamespaceNode::Type {
        ctor: Some(CtorSpec {
            ancestry: vec![name.to_string()],
            ..Default::default()
        }),
        ctor_call: Some(Box::new(NamespaceNode::CommonEmit(ctor_emit.to_string()))),
        statics: Subtree::new(),
        methods: methods
            .iter()
            .map(|(method, emit)| {
                (
                    (*method).to_string(),
                    NamespaceNode::CommonEmit((*emit).to_string()),
                )
            })
            .collect(),
        member_returns: BTreeMap::new(),
    }
}

fn doctest_subtree() -> Subtree {
    let mut root = Subtree::new();
    for (name, value) in [
        ("DONT_ACCEPT_TRUE_FOR_1", 1 << 0),
        ("DONT_ACCEPT_BLANKLINE", 1 << 1),
        ("NORMALIZE_WHITESPACE", 1 << 2),
        ("ELLIPSIS", 1 << 3),
        ("SKIP", 1 << 4),
        ("IGNORE_EXCEPTION_DETAIL", 1 << 5),
        ("REPORT_UDIFF", 1 << 6),
        ("REPORT_CDIFF", 1 << 7),
        ("REPORT_NDIFF", 1 << 8),
    ] {
        root.insert(name.to_string(), NamespaceNode::Const(Value::I32(value)));
    }
    root.insert(
        "DocTestParser".to_string(),
        simple_python_type(
            "DocTestParser",
            "python.doctest_parser_new",
            &[
                ("get_examples", "python.doctest_parser_get_examples"),
                ("get_doctest", "python.doctest_parser_get_doctest"),
            ],
        ),
    );
    root.insert(
        "DocTestFinder".to_string(),
        simple_python_type(
            "DocTestFinder",
            "python.doctest_finder_new",
            &[("find", "python.doctest_finder_find")],
        ),
    );
    root.insert(
        "DocTestRunner".to_string(),
        simple_python_type(
            "DocTestRunner",
            "python.doctest_runner_new",
            &[
                ("run", "python.doctest_runner_run"),
                ("summarize", "python.doctest_runner_summarize"),
            ],
        ),
    );
    root.insert(
        "OutputChecker".to_string(),
        simple_python_type(
            "OutputChecker",
            "python.doctest_output_checker_new",
            &[
                ("check_output", "python.doctest_output_checker_check_output"),
                (
                    "output_difference",
                    "python.doctest_output_checker_output_difference",
                ),
            ],
        ),
    );
    root.insert(
        "Example".to_string(),
        NamespaceNode::CommonEmit("python.doctest_example".to_string()),
    );
    root.insert(
        "register_optionflag".to_string(),
        NamespaceNode::CommonEmit("python.doctest_register_optionflag".to_string()),
    );
    root.insert(
        "script_from_examples".to_string(),
        NamespaceNode::CommonEmit("python.doctest_script_from_examples".to_string()),
    );
    root.insert(
        "testmod".to_string(),
        NamespaceNode::CommonEmit("python.doctest_testmod".to_string()),
    );
    root
}

fn pydoc_subtree() -> Subtree {
    let mut root = Subtree::new();
    for name in ["Helper", "TextDoc", "HTMLDoc"] {
        root.insert(
            name.to_string(),
            simple_python_type(
                name,
                &format!("python.pydoc_{name}_new"),
                &[("document", "python.pydoc_document"), ("render_doc", "python.pydoc_render_doc")],
            ),
        );
    }
    for (name, emit) in [
        ("stripid", "python.pydoc_stripid"),
        ("splitdoc", "python.pydoc_splitdoc"),
        ("classname", "python.pydoc_classname"),
        ("describe", "python.pydoc_describe"),
        ("locate", "python.pydoc_locate"),
        ("resolve", "python.pydoc_resolve"),
        ("render_doc", "python.pydoc_render_doc"),
        ("allmethods", "python.pydoc_allmethods"),
    ] {
        root.insert(name.to_string(), NamespaceNode::CommonEmit(emit.to_string()));
    }
    root
}

fn dis_subtree() -> Subtree {
    let mut root = Subtree::new();
    for (name, emit) in [
        ("Bytecode", "python.dis_bytecode"),
        ("code_info", "python.dis_code_info"),
        ("dis", "python.dis_dis"),
        ("disassemble", "python.dis_disassemble"),
        ("show_code", "python.dis_show_code"),
        ("get_instructions", "python.dis_get_instructions"),
        ("findlabels", "python.dis_findlabels"),
        ("findlinestarts", "python.dis_findlinestarts"),
        ("stack_effect", "python.dis_stack_effect"),
    ] {
        root.insert(name.to_string(), NamespaceNode::CommonEmit(emit.to_string()));
    }
    root.insert(
        "Instruction".to_string(),
        NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics: Subtree::new(),
            methods: Subtree::new(),
            member_returns: BTreeMap::new(),
        },
    );
    root.insert(
        "opmap".to_string(),
        NamespaceNode::CommonEmit("python.dis_opmap".to_string()),
    );
    root.insert(
        "opname".to_string(),
        NamespaceNode::CommonEmit("python.dis_opname".to_string()),
    );
    root.insert(
        "cmp_op".to_string(),
        NamespaceNode::Const(Value::String(std::sync::Arc::from(
            "< <= == != > >=",
        ))),
    );
    root.insert(
        "hasconst".to_string(),
        NamespaceNode::CommonEmit("python.dis_hasconst".to_string()),
    );
    root.insert(
        "hasname".to_string(),
        NamespaceNode::CommonEmit("python.dis_hasname".to_string()),
    );
    root.insert(
        "haslocal".to_string(),
        NamespaceNode::CommonEmit("python.dis_haslocal".to_string()),
    );
    root
}

fn symtable_subtree() -> Subtree {
    let mut root = Subtree::new();
    root.insert(
        "symtable".to_string(),
        NamespaceNode::CommonEmit("python.symtable_symtable".to_string()),
    );
    for name in ["Symbol", "SymbolTable", "Function", "Class"] {
        root.insert(
            name.to_string(),
            NamespaceNode::Type {
                ctor: None,
                ctor_call: None,
                statics: Subtree::new(),
                methods: Subtree::new(),
                member_returns: BTreeMap::new(),
            },
        );
    }
    for (name, value) in [
        ("TYPE_MODULE", "module"),
        ("TYPE_FUNCTION", "function"),
        ("TYPE_CLASS", "class"),
    ] {
        root.insert(
            name.to_string(),
            NamespaceNode::Const(Value::String(std::sync::Arc::from(value))),
        );
    }
    root
}

fn token_subtree(include_tokenize_only: bool) -> Subtree {
    let mut root = Subtree::new();
    for (name, value) in [
        ("ENDMARKER", 0),
        ("NAME", 1),
        ("NUMBER", 2),
        ("STRING", 3),
        ("NEWLINE", 4),
        ("INDENT", 5),
        ("DEDENT", 6),
        ("OP", 54),
        ("COMMENT", 61),
        ("NL", 62),
        ("ENCODING", 63),
    ] {
        if include_tokenize_only || name != "COMMENT" {
            root.insert(name.to_string(), NamespaceNode::Const(Value::I32(value)));
        }
    }
    if include_tokenize_only {
        for name in ["TokenInfo", "TokenError"] {
            root.insert(
                name.to_string(),
                NamespaceNode::Type {
                    ctor: None,
                    ctor_call: None,
                    statics: Subtree::new(),
                    methods: Subtree::new(),
                    member_returns: BTreeMap::new(),
                },
            );
        }
    }
    root
}

fn email_subtree() -> Subtree {
    let mut root = Subtree::new();
    root.insert(
        "message_from_string".to_string(),
        NamespaceNode::CommonEmit("python.email_message_from_string".to_string()),
    );
    root.insert(
        "message_from_bytes".to_string(),
        NamespaceNode::CommonEmit("python.email_message_from_bytes".to_string()),
    );

    let mut message = Subtree::new();
    message.insert(
        "EmailMessage".to_string(),
        simple_python_type(
            "EmailMessage",
            "python.email_message_new",
            &[
                ("set_content", "python.email_message_set_content"),
                ("add_header", "python.email_message_add_header"),
                ("replace_header", "python.email_message_replace_header"),
                ("get_all", "python.email_message_get_all"),
                ("get_content_type", "python.email_message_get_content_type"),
                ("get_content", "python.email_message_get_content"),
                ("get_payload", "python.email_message_get_payload"),
                ("is_multipart", "python.email_message_is_multipart"),
                ("iter_parts", "python.email_message_iter_parts"),
                ("walk", "python.email_message_walk"),
                ("as_string", "python.email_message_as_string"),
                ("as_bytes", "python.email_message_as_bytes"),
                ("add_alternative", "python.email_message_add_alternative"),
                ("add_attachment", "python.email_message_add_attachment"),
            ],
        ),
    );
    root.insert("message".to_string(), NamespaceNode::Namespace(message));

    let mut header = Subtree::new();
    header.insert(
        "decode_header".to_string(),
        NamespaceNode::CommonEmit("python.email_decode_header".to_string()),
    );
    header.insert(
        "make_header".to_string(),
        NamespaceNode::CommonEmit("python.email_make_header".to_string()),
    );
    root.insert("header".to_string(), NamespaceNode::Namespace(header));

    let mut utils = Subtree::new();
    for (name, emit) in [
        ("parseaddr", "python.email_parseaddr"),
        ("formataddr", "python.email_formataddr"),
        ("formatdate", "python.email_formatdate"),
        ("parsedate_to_datetime", "python.email_parsedate_to_datetime"),
    ] {
        utils.insert(name.to_string(), NamespaceNode::CommonEmit(emit.to_string()));
    }
    root.insert("utils".to_string(), NamespaceNode::Namespace(utils));

    let mut policy = Subtree::new();
    policy.insert(
        "default".to_string(),
        NamespaceNode::Const(Value::String(std::sync::Arc::from("default"))),
    );
    root.insert("policy".to_string(), NamespaceNode::Namespace(policy));
    root
}

/// The classes `core_classes` declares, as type nodes under their module.
///
/// The name list comes from `core_classes::MODULE_SURFACE`, so the tree cannot
/// declare a type the AST does not build, or miss one it does — the same
/// coupling dart's `core_types()` keeps.
///
/// ⛔ These are NOT construction leaves. A declared class is compiled into the
/// module as an ordinary global, so `ipaddress.IPv4Network(x)` constructs
/// through that global; the tree node is what the MEMBER-READ path and the
/// static-type hint resolve against (`expressions.rs` reads a registered type's
/// `methods`/`member_returns`), and what makes the module's surface
/// enumerable. Receiver dispatch never comes here — that is the prototype the
/// class compiler stamps.
fn core_class_subtrees() -> Vec<(String, Subtree)> {
    let mut by_module: BTreeMap<String, Subtree> = BTreeMap::new();
    for (module, name, global) in crate::core_classes::MODULE_SURFACE {
        by_module.entry((*module).to_string()).or_default().insert(
            (*name).to_string(),
            NamespaceNode::Type {
                ctor: None,
                ctor_call: None,
                statics: Subtree::new(),
                methods: Subtree::new(),
                member_returns: BTreeMap::from([(
                    "__vybe_declared_global".to_string(),
                    (*global).to_string(),
                )]),
            },
        );
    }
    by_module.into_iter().collect()
}
