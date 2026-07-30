//! Java walker — pest `Pair<Rule>` → `vybe_compiler::ast::Module`.
//!
//! Walks the parse tree produced by `grammar.pest` into the common AST.
//! Once this returns a `Module`, the rest of the compilation pipeline
//! is shared with every other vybex language.
//!
//! ## Walker normalisations
//!
//! - **`System.out.println` / `System.out.print`**: rewritten in the walker
//!   to bare `println` / `print` calls so the profile can bind them.
//! - **`System.exit(code)`**: rewritten to `__process_exit(code)`.
//! - **Enhanced-for (`for (T x : iterable)`)**: `ForIn { of: true }` — Java
//!   always iterates values.
//! - **Implicit `super()` in child-class constructors**: injected when
//!   no explicit `super(...)` is found at the top of the ctor body.
//! - **Generic type arguments**: parsed through the shared generics primitive,
//!   then erased for runtime dispatch while preserved in type hints/metadata.
//! - **Char literals**: lowered to integer code-point literals.
//! - **Text blocks** (Java 15+): normalised to plain string literals.
//! - **Lambda params**: both typed `(T x) -> body` and untyped `x -> body`
//!   forms reduced to bare name params.

use super::{JavaParser, Rule};
use pest::Parser;
use pest::iterators::Pair;
use std::cell::RefCell;
use std::collections::{HashMap, HashSet};
use vybe_ast::*;
use vybe_compiler::primitives::generics as common_generics;
use vybe_compiler::primitives::reflection as common_reflection;

#[derive(Clone, Debug, Default)]
struct JavaReflectionClassMeta {
    parent: Option<String>,
    interfaces: Vec<String>,
    fields: Vec<JavaReflectionFieldMeta>,
    methods: Vec<JavaReflectionCallableMeta>,
    constructors: Vec<JavaReflectionCallableMeta>,
    nested_classes: Vec<String>,
    modifiers: i64,
    is_interface: bool,
    is_enum: bool,
}

#[derive(Clone, Debug, Default)]
struct JavaReflectionCallableMeta {
    name: String,
    param_count: usize,
    param_types: Vec<String>,
    return_type: Option<String>,
    modifiers: i64,
}

#[derive(Clone, Debug, Default)]
struct JavaReflectionFieldMeta {
    name: String,
    type_name: Option<String>,
    modifiers: i64,
}

const JAVA_MOD_PUBLIC: i64 = 0x0001;
const JAVA_MOD_PRIVATE: i64 = 0x0002;
const JAVA_MOD_PROTECTED: i64 = 0x0004;
const JAVA_MOD_STATIC: i64 = 0x0008;
const JAVA_MOD_FINAL: i64 = 0x0010;
const JAVA_MOD_ABSTRACT: i64 = 0x0400;

thread_local! {
    static JAVA_INTERFACE_NAMES: RefCell<HashSet<String>> = RefCell::new(HashSet::new());
    static JAVA_NESTED_TYPE_NAMES: RefCell<HashSet<String>> = RefCell::new(HashSet::new());
    static JAVA_ENUM_VALUES: RefCell<HashMap<String, Vec<String>>> = RefCell::new(HashMap::new());
    static JAVA_RECORD_COMPONENTS: RefCell<HashMap<String, HashSet<String>>> = RefCell::new(HashMap::new());
    static JAVA_REFLECTION_CLASSES: RefCell<HashMap<String, JavaReflectionClassMeta>> = RefCell::new(HashMap::new());
    static JAVA_CURRENT_CLASS_STACK: RefCell<Vec<String>> = RefCell::new(Vec::new());
    // Locals declared with a PrintStream type — their method calls route
    // through the __j_* PrintStream runtime like `System.out` itself.
    static JAVA_PRINTSTREAM_VARS: RefCell<HashSet<String>> = RefCell::new(HashSet::new());
    // Locals declared as StringBuilder/StringBuffer — their method calls
    // route through the __j_sb_* runtime (emitter/format_runtime.rs).
    static JAVA_SB_VARS: RefCell<HashSet<String>> = RefCell::new(HashSet::new());
    static JAVA_STRING_VARS: RefCell<HashSet<String>> = RefCell::new(HashSet::new());
    static JAVA_DOUBLE_VARS: RefCell<HashSet<String>> = RefCell::new(HashSet::new());
    static JAVA_STRING_JOINER_VARS: RefCell<HashSet<String>> = RefCell::new(HashSet::new());
    static JAVA_STRING_TOKENIZER_VARS: RefCell<HashSet<String>> = RefCell::new(HashSet::new());
    static JAVA_SCANNER_VARS: RefCell<HashSet<String>> = RefCell::new(HashSet::new());
    static JAVA_FORMATTER_VARS: RefCell<HashSet<String>> = RefCell::new(HashSet::new());
    static JAVA_MESSAGE_FORMAT_VARS: RefCell<HashSet<String>> = RefCell::new(HashSet::new());
    static JAVA_CALENDAR_VARS: RefCell<HashSet<String>> = RefCell::new(HashSet::new());
    static JAVA_THREAD_VARS: RefCell<HashSet<String>> = RefCell::new(HashSet::new());
    static JAVA_THREAD_TYPES: RefCell<HashSet<String>> = RefCell::new(HashSet::new());
    static JAVA_THREAD_TARGETS: RefCell<HashMap<String, Expression>> = RefCell::new(HashMap::new());
    static JAVA_RUNNABLE_TARGETS: RefCell<HashMap<String, Expression>> = RefCell::new(HashMap::new());
    static JAVA_THREAD_UNSAFE_TARGETS: RefCell<HashSet<String>> = RefCell::new(HashSet::new());
    static JAVA_RUNNABLE_UNSAFE_TARGETS: RefCell<HashSet<String>> = RefCell::new(HashSet::new());
    static JAVA_STATIC_FIELD_VARS: RefCell<HashSet<String>> = RefCell::new(HashSet::new());
    static JAVA_STATIC_FIELD_TYPES: RefCell<HashMap<String, String>> = RefCell::new(HashMap::new());
    static JAVA_LOCAL_TYPES: RefCell<HashMap<String, String>> = RefCell::new(HashMap::new());
    static JAVA_FINAL_CONSTANTS: RefCell<HashMap<String, Expression>> = RefCell::new(HashMap::new());
    static JAVA_RUNNABLE_VARS: RefCell<HashSet<String>> = RefCell::new(HashSet::new());
    static JAVA_FUNCTIONAL_VARS: RefCell<HashSet<String>> = RefCell::new(HashSet::new());
    static JAVA_FUNCTIONAL_TYPES: RefCell<HashMap<String, String>> = RefCell::new(HashMap::new());
    static JAVA_FUNCTIONAL_INTERFACE_METHODS: RefCell<HashMap<String, String>> = RefCell::new(HashMap::new());
    static JAVA_OPTIONAL_VARS: RefCell<HashSet<String>> = RefCell::new(HashSet::new());
    static JAVA_RANDOM_VARS: RefCell<HashSet<String>> = RefCell::new(HashSet::new());
    static JAVA_TLR_VARS: RefCell<HashSet<String>> = RefCell::new(HashSet::new());
    static JAVA_CHAR_ARRAY_VARS: RefCell<HashSet<String>> = RefCell::new(HashSet::new());
    static JAVA_BYTE_ARRAY_VARS: RefCell<HashSet<String>> = RefCell::new(HashSet::new());
    static JAVA_BIGINT_VARS: RefCell<HashSet<String>> = RefCell::new(HashSet::new());
    static JAVA_BIGDECIMAL_VARS: RefCell<HashSet<String>> = RefCell::new(HashSet::new());
    static JAVA_DECIMAL_FORMAT_VARS: RefCell<HashSet<String>> = RefCell::new(HashSet::new());
    static JAVA_NUMBER_VARS: RefCell<HashSet<String>> = RefCell::new(HashSet::new());
    // Locals declared as java.util.regex Pattern / Matcher — routed
    // through the __j_pat_*/__j_m_* runtime.
    static JAVA_PATTERN_VARS: RefCell<HashSet<String>> = RefCell::new(HashSet::new());
    static JAVA_MATCHER_VARS: RefCell<HashSet<String>> = RefCell::new(HashSet::new());
    // Pending `instanceof T var` pattern bindings (JLS §14.30.1): each is
    // (binding var, type name, subject expr). walk_if drains the bindings
    // its condition produced and prepends `T var = subject;` to the
    // then-body (flow-scoped, like the record/switch pattern paths).
    static JAVA_INSTANCEOF_BINDINGS: RefCell<Vec<(String, String, Expression)>> =
        RefCell::new(Vec::new());
    // Locals declared as java.net.URL/URI — __j_url_* runtime.
    static JAVA_URL_VARS: RefCell<HashSet<String>> = RefCell::new(HashSet::new());
    // Locals declared as javax.xml.namespace.QName — common XML name runtime.
    static JAVA_QNAME_VARS: RefCell<HashSet<String>> = RefCell::new(HashSet::new());
}

// ════════════════════════════════════════════════════════════════════════════
// Entry point
// ════════════════════════════════════════════════════════════════════════════

pub fn parse(source: &str) -> Result<Module, String> {
    JAVA_INTERFACE_NAMES.with(|names| names.borrow_mut().clear());
    JAVA_NESTED_TYPE_NAMES.with(|names| names.borrow_mut().clear());
    JAVA_ENUM_VALUES.with(|values| values.borrow_mut().clear());
    JAVA_RECORD_COMPONENTS.with(|components| components.borrow_mut().clear());
    JAVA_REFLECTION_CLASSES.with(|classes| classes.borrow_mut().clear());
    JAVA_CURRENT_CLASS_STACK.with(|stack| stack.borrow_mut().clear());
    JAVA_PRINTSTREAM_VARS.with(|vars| vars.borrow_mut().clear());
    JAVA_SB_VARS.with(|vars| vars.borrow_mut().clear());
    JAVA_STRING_VARS.with(|vars| vars.borrow_mut().clear());
    JAVA_DOUBLE_VARS.with(|vars| vars.borrow_mut().clear());
    JAVA_STRING_JOINER_VARS.with(|vars| vars.borrow_mut().clear());
    JAVA_STRING_TOKENIZER_VARS.with(|vars| vars.borrow_mut().clear());
    JAVA_SCANNER_VARS.with(|vars| vars.borrow_mut().clear());
    JAVA_FORMATTER_VARS.with(|vars| vars.borrow_mut().clear());
    JAVA_MESSAGE_FORMAT_VARS.with(|vars| vars.borrow_mut().clear());
    JAVA_CALENDAR_VARS.with(|vars| vars.borrow_mut().clear());
    JAVA_THREAD_VARS.with(|vars| vars.borrow_mut().clear());
    JAVA_THREAD_TYPES.with(|vars| vars.borrow_mut().clear());
    JAVA_THREAD_TARGETS.with(|targets| targets.borrow_mut().clear());
    JAVA_RUNNABLE_TARGETS.with(|targets| targets.borrow_mut().clear());
    JAVA_THREAD_UNSAFE_TARGETS.with(|targets| targets.borrow_mut().clear());
    JAVA_RUNNABLE_UNSAFE_TARGETS.with(|targets| targets.borrow_mut().clear());
    JAVA_STATIC_FIELD_VARS.with(|vars| vars.borrow_mut().clear());
    JAVA_STATIC_FIELD_TYPES.with(|types| types.borrow_mut().clear());
    JAVA_LOCAL_TYPES.with(|types| types.borrow_mut().clear());
    JAVA_FINAL_CONSTANTS.with(|constants| constants.borrow_mut().clear());
    JAVA_RUNNABLE_VARS.with(|vars| vars.borrow_mut().clear());
    JAVA_FUNCTIONAL_VARS.with(|vars| vars.borrow_mut().clear());
    JAVA_FUNCTIONAL_TYPES.with(|types| types.borrow_mut().clear());
    JAVA_FUNCTIONAL_INTERFACE_METHODS.with(|methods| methods.borrow_mut().clear());
    JAVA_OPTIONAL_VARS.with(|vars| vars.borrow_mut().clear());
    JAVA_RANDOM_VARS.with(|vars| vars.borrow_mut().clear());
    JAVA_TLR_VARS.with(|vars| vars.borrow_mut().clear());
    JAVA_CHAR_ARRAY_VARS.with(|vars| vars.borrow_mut().clear());
    JAVA_BYTE_ARRAY_VARS.with(|vars| vars.borrow_mut().clear());
    JAVA_BIGINT_VARS.with(|vars| vars.borrow_mut().clear());
    JAVA_BIGDECIMAL_VARS.with(|vars| vars.borrow_mut().clear());
    JAVA_DECIMAL_FORMAT_VARS.with(|vars| vars.borrow_mut().clear());
    JAVA_NUMBER_VARS.with(|vars| vars.borrow_mut().clear());
    JAVA_PATTERN_VARS.with(|vars| vars.borrow_mut().clear());
    JAVA_MATCHER_VARS.with(|vars| vars.borrow_mut().clear());
    JAVA_URL_VARS.with(|vars| vars.borrow_mut().clear());
    JAVA_QNAME_VARS.with(|vars| vars.borrow_mut().clear());
    JAVA_INSTANCEOF_BINDINGS.with(|b| b.borrow_mut().clear());

    let mut pairs =
        JavaParser::parse(Rule::program, source).map_err(|e| format!("Java parse error: {}", e))?;
    let program = pairs.next().ok_or("empty parse")?;

    let mut body = Vec::new();
    let mut imports = Vec::new();
    // JLS §7.5.3/§7.5.4 static imports, resolved by walker rewrite:
    // member name → declaring type; on-demand list of declaring types.
    let mut static_single_members: HashMap<String, String> = HashMap::new();
    let mut static_on_demand_types: Vec<String> = Vec::new();

    for pair in program.into_inner() {
        match pair.as_rule() {
            Rule::EOI => continue,
            Rule::package_declaration => {}
            Rule::import_declaration => {
                let text = pair
                    .as_str()
                    .trim_start_matches("import")
                    .trim()
                    .trim_end_matches(';')
                    .trim();
                if let Some(member_path) = text.strip_prefix("static") {
                    // JLS §7.5.3/§7.5.4 static imports: fully consumed by
                    // the walker rewrite below (bare `max(…)` →
                    // `Math.max(…)`), recorded as an inert Simple import.
                    let member_path = member_path.trim();
                    let mut segments: Vec<&str> = member_path.split('.').collect();
                    if let (Some(last), Some(ty)) = (segments.pop(), segments.last().copied()) {
                        if last == "*" {
                            static_on_demand_types.push(ty.to_string());
                        } else if segments.len() >= 1 {
                            static_single_members.insert(last.to_string(), ty.to_string());
                        }
                    }
                    imports.push(Import {
                        kind: ImportKind::Simple {
                            path: member_path.to_string(),
                            alias: None,
                        },
                        span: to_span(&pair),
                    });
                } else if let Some(imp) = walk_import(&pair) {
                    imports.push(imp);
                }
            }
            Rule::class_declaration => body.push(Statement::new(walk_class(pair)?)),
            Rule::interface_declaration => body.push(Statement::new(walk_interface(pair)?)),
            Rule::enum_declaration => body.push(Statement::new(walk_enum_decl(pair)?)),
            Rule::record_declaration => body.push(Statement::new(walk_record(pair)?)),
            Rule::annotation_type_declaration => {}
            _ => {
                if let Some(s) = walk_statement(pair)? {
                    body.push(s);
                }
            }
        }
    }

    rewrite_java_static_imports(&mut body, &static_single_members, &static_on_demand_types);
    // Local (method-body) classes → static nested siblings of the enclosing
    // class, with any captured enclosing locals threaded through the
    // constructor. Runs before nested-type qualification so hoisted classes
    // are treated as ordinary nested types by every pass below.
    hoist_java_local_classes(&mut body);
    qualify_java_nested_types(&mut body);
    rewrite_java_user_tostring_calls(&mut body);
    erase_java_interface_param_hints(&mut body);
    reject_java_direct_abstract_instantiation(&body)?;

    // PrintStream/Formatter runtime (emitter/format_runtime.rs). Keep it
    // available for the composite runtime pieces that have not moved to
    // profile/common emitters yet, but do not prepend it for modules that only
    // use direct emitter-backed helpers such as basic print/println.
    let prelude = super::emitter::format_runtime::prelude();
    let mut body = if java_body_references_prelude(&body, &prelude) {
        let mut with_prelude = prelude;
        with_prelude.append(&mut body);
        with_prelude
    } else {
        body
    };
    normalize_java_class_tree(&mut body);
    strip_java_abstract_method_declarations(&mut body);
    lower_java_abstract_runtime_modifiers(&mut body);
    inject_java_static_initializer_calls(&mut body);

    // Java: inject a top-level call to the class's static main method.
    // Uses the same pattern as EntryPoint::Method in bundle.rs.
    if let Some(class_name) = find_main_class(&body) {
        body.push(Statement::new(StmtKind::Expr(Expression::new(
            ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(Expression::ident(&class_name)),
                    field: "main".to_string(),
                    null_safe: false,
                })),
                args: vec![],
                optional: false,
            },
        ))));
    }

    Ok(Module {
        name: String::new(),
        language: Lang::Java,
        body,
        imports,
    })
}

fn java_body_references_prelude(body: &[Statement], prelude: &[Statement]) -> bool {
    let names = java_prelude_top_level_names(prelude);
    if names.is_empty() {
        return false;
    }
    body.iter()
        .any(|stmt| java_stmt_references_any_name(stmt, &names))
}

fn java_prelude_top_level_names(prelude: &[Statement]) -> HashSet<String> {
    let mut names = HashSet::new();
    for stmt in prelude {
        match &stmt.kind {
            StmtKind::FunctionDecl { name, .. } | StmtKind::ClassDecl { name, .. } => {
                names.insert(name.clone());
            }
            StmtKind::VarDecl { declarations, .. } => {
                for decl in declarations {
                    if let BindingPattern::Ident(name) = &decl.pattern {
                        names.insert(name.clone());
                    }
                }
            }
            _ => {}
        }
    }
    names
}

fn java_stmt_references_any_name(stmt: &Statement, names: &HashSet<String>) -> bool {
    match &stmt.kind {
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            java_expr_references_any_name(expr, names)
        }
        StmtKind::Block(body) | StmtKind::FunctionDecl { body, .. } => body
            .iter()
            .any(|stmt| java_stmt_references_any_name(stmt, names)),
        StmtKind::ClassDecl {
            parents,
            interfaces,
            members,
            decorators,
            ..
        } => {
            parents.iter().any(|name| names.contains(name))
                || interfaces.iter().any(|name| names.contains(name))
                || decorators
                    .iter()
                    .any(|expr| java_expr_references_any_name(expr, names))
                || members
                    .iter()
                    .any(|member| java_class_member_references_any_name(member, names))
        }
        StmtKind::InterfaceDecl {
            parents,
            decorators,
            ..
        } => {
            parents.iter().any(|name| names.contains(name))
                || decorators
                    .iter()
                    .any(|expr| java_expr_references_any_name(expr, names))
        }
        StmtKind::EnumDecl {
            interfaces,
            body_members,
            decorators,
            ..
        } => {
            interfaces.iter().any(|name| names.contains(name))
                || decorators
                    .iter()
                    .any(|expr| java_expr_references_any_name(expr, names))
                || body_members
                    .iter()
                    .any(|member| java_class_member_references_any_name(member, names))
        }
        StmtKind::StructDecl {
            interfaces,
            members,
            decorators,
            ..
        } => {
            interfaces.iter().any(|name| names.contains(name))
                || decorators
                    .iter()
                    .any(|expr| java_expr_references_any_name(expr, names))
                || members
                    .iter()
                    .any(|member| java_class_member_references_any_name(member, names))
        }
        StmtKind::NamespaceDecl { body, .. } => body
            .iter()
            .any(|stmt| java_stmt_references_any_name(stmt, names)),
        StmtKind::VarDecl { declarations, .. } => declarations.iter().any(|decl| {
            decl.type_hint
                .as_deref()
                .is_some_and(|type_name| names.contains(type_name))
                || decl
                    .init
                    .as_ref()
                    .is_some_and(|expr| java_expr_references_any_name(expr, names))
        }),
        StmtKind::Assign { targets, value } => {
            targets
                .iter()
                .any(|expr| java_expr_references_any_name(expr, names))
                || java_expr_references_any_name(value, names)
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            java_expr_references_any_name(target, names)
                || java_expr_references_any_name(value, names)
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            java_expr_references_any_name(cond, names)
                || then_body
                    .iter()
                    .any(|stmt| java_stmt_references_any_name(stmt, names))
                || elifs.iter().any(|(cond, body)| {
                    java_expr_references_any_name(cond, names)
                        || body
                            .iter()
                            .any(|stmt| java_stmt_references_any_name(stmt, names))
                })
                || else_body.as_ref().is_some_and(|body| {
                    body.iter()
                        .any(|stmt| java_stmt_references_any_name(stmt, names))
                })
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            init.as_ref()
                .is_some_and(|stmt| java_stmt_references_any_name(stmt, names))
                || cond
                    .as_ref()
                    .is_some_and(|expr| java_expr_references_any_name(expr, names))
                || update
                    .as_ref()
                    .is_some_and(|expr| java_expr_references_any_name(expr, names))
                || body
                    .iter()
                    .any(|stmt| java_stmt_references_any_name(stmt, names))
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            java_expr_references_any_name(iter, names)
                || body
                    .iter()
                    .any(|stmt| java_stmt_references_any_name(stmt, names))
                || else_body.as_ref().is_some_and(|body| {
                    body.iter()
                        .any(|stmt| java_stmt_references_any_name(stmt, names))
                })
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            java_expr_references_any_name(cond, names)
                || body
                    .iter()
                    .any(|stmt| java_stmt_references_any_name(stmt, names))
                || else_body.as_ref().is_some_and(|body| {
                    body.iter()
                        .any(|stmt| java_stmt_references_any_name(stmt, names))
                })
        }
        StmtKind::DoWhile { body, cond, .. } => {
            body.iter()
                .any(|stmt| java_stmt_references_any_name(stmt, names))
                || java_expr_references_any_name(cond, names)
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            java_expr_references_any_name(expr, names)
                || cases.iter().any(|case| {
                    case.conditions
                        .iter()
                        .any(|condition| java_case_condition_references_any_name(condition, names))
                        || case
                            .body
                            .iter()
                            .any(|stmt| java_stmt_references_any_name(stmt, names))
                })
                || default.as_ref().is_some_and(|body| {
                    body.iter()
                        .any(|stmt| java_stmt_references_any_name(stmt, names))
                })
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
            ..
        } => {
            body.iter()
                .any(|stmt| java_stmt_references_any_name(stmt, names))
                || catches.iter().any(|catch| {
                    catch
                        .body
                        .iter()
                        .any(|stmt| java_stmt_references_any_name(stmt, names))
                })
                || else_body.as_ref().is_some_and(|body| {
                    body.iter()
                        .any(|stmt| java_stmt_references_any_name(stmt, names))
                })
                || finally.as_ref().is_some_and(|body| {
                    body.iter()
                        .any(|stmt| java_stmt_references_any_name(stmt, names))
                })
        }
        StmtKind::Throw { expr, cause } => {
            expr.as_ref()
                .is_some_and(|expr| java_expr_references_any_name(expr, names))
                || cause
                    .as_ref()
                    .is_some_and(|expr| java_expr_references_any_name(expr, names))
        }
        _ => false,
    }
}

fn java_class_member_references_any_name(member: &ClassMember, names: &HashSet<String>) -> bool {
    match member {
        ClassMember::Field {
            type_hint,
            init,
            array_bounds,
            ..
        } => {
            type_hint
                .as_deref()
                .is_some_and(|type_name| names.contains(type_name))
                || init
                    .as_ref()
                    .is_some_and(|expr| java_expr_references_any_name(expr, names))
                || array_bounds.as_ref().is_some_and(|bounds| {
                    bounds
                        .iter()
                        .any(|expr| java_expr_references_any_name(expr, names))
                })
        }
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            java_stmt_references_any_name(stmt, names)
        }
        ClassMember::Constructor {
            body, base_args, ..
        } => {
            base_args.as_ref().is_some_and(|args| {
                args.iter()
                    .any(|expr| java_expr_references_any_name(expr, names))
            }) || body
                .iter()
                .any(|stmt| java_stmt_references_any_name(stmt, names))
        }
        ClassMember::Property {
            type_hint,
            getter,
            setter,
            ..
        } => {
            type_hint
                .as_deref()
                .is_some_and(|type_name| names.contains(type_name))
                || getter.as_ref().is_some_and(|body| {
                    body.iter()
                        .any(|stmt| java_stmt_references_any_name(stmt, names))
                })
                || setter.as_ref().is_some_and(|setter| {
                    setter
                        .body
                        .iter()
                        .any(|stmt| java_stmt_references_any_name(stmt, names))
                })
        }
        ClassMember::Const {
            type_hint, value, ..
        } => {
            type_hint
                .as_deref()
                .is_some_and(|type_name| names.contains(type_name))
                || java_expr_references_any_name(value, names)
        }
        ClassMember::Event { type_hint, .. } => type_hint
            .as_deref()
            .is_some_and(|type_name| names.contains(type_name)),
        // An augmentation names the type it draws from — that IS a reference to
        // it, and treating it as none would let a type look unused.
        ClassMember::Augment(decl) => names.contains(&decl.from),
    }
}

fn java_case_condition_references_any_name(
    condition: &CaseCondition,
    names: &HashSet<String>,
) -> bool {
    match condition {
        CaseCondition::Value(expr) | CaseCondition::Comparison { expr, .. } => {
            java_expr_references_any_name(expr, names)
        }
        CaseCondition::Range { from, to } => {
            java_expr_references_any_name(from, names) || java_expr_references_any_name(to, names)
        }
    }
}

fn java_expr_references_any_name(expr: &Expression, names: &HashSet<String>) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => names.contains(name),
        ExprKind::Call { callee, args, .. } => {
            java_expr_references_any_name(callee, names)
                || args
                    .iter()
                    .any(|arg| java_expr_references_any_name(&arg.value, names))
        }
        ExprKind::New { class, args } => {
            java_expr_references_any_name(class, names)
                || args
                    .iter()
                    .any(|arg| java_expr_references_any_name(&arg.value, names))
        }
        ExprKind::Member { object, .. } => java_expr_references_any_name(object, names),
        ExprKind::Index { object, index, .. } => {
            java_expr_references_any_name(object, names)
                || java_expr_references_any_name(index, names)
        }
        ExprKind::Binary { left, right, .. } => {
            java_expr_references_any_name(left, names)
                || java_expr_references_any_name(right, names)
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::Spread(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr)
        | ExprKind::RefLoad(expr) => java_expr_references_any_name(expr, names),
        ExprKind::Yield(Some(expr)) => java_expr_references_any_name(expr, names),
        ExprKind::Assign { target, value } => {
            java_expr_references_any_name(target, names)
                || java_expr_references_any_name(value, names)
        }
        ExprKind::Ternary { cond, then, else_ } => {
            java_expr_references_any_name(cond, names)
                || java_expr_references_any_name(then, names)
                || java_expr_references_any_name(else_, names)
        }
        ExprKind::Array(elems) => elems.iter().any(|elem| {
            elem.key
                .as_ref()
                .is_some_and(|expr| java_expr_references_any_name(expr, names))
                || java_expr_references_any_name(&elem.value, names)
        }),
        ExprKind::Tuple(items) | ExprKind::Set(items) | ExprKind::Sequence(items) => items
            .iter()
            .any(|expr| java_expr_references_any_name(expr, names)),
        ExprKind::Object(props) => props.iter().any(|prop| match prop {
            ObjectProperty::KeyValue { key, value } | ObjectProperty::Computed { key, value } => {
                java_expr_references_any_name(key, names)
                    || java_expr_references_any_name(value, names)
            }
            ObjectProperty::Spread(expr) => java_expr_references_any_name(expr, names),
            ObjectProperty::Method { value, .. } | ObjectProperty::Accessor { value, .. } => {
                java_stmt_references_any_name(value, names)
            }
            ObjectProperty::Shorthand(name) => names.contains(name),
        }),
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(expr) => java_expr_references_any_name(expr, names),
            LambdaBody::Block(body) => body
                .iter()
                .any(|stmt| java_stmt_references_any_name(stmt, names)),
        },
        ExprKind::NamedTuple { fields, .. } => fields
            .iter()
            .any(|(_, expr)| java_expr_references_any_name(expr, names)),
        ExprKind::Interpolation(parts) => parts.iter().any(|part| match part {
            InterpolPart::Expr(expr) => java_expr_references_any_name(expr, names),
            InterpolPart::Formatted(expr, _) => java_expr_references_any_name(expr, names),
            InterpolPart::Text(_) => false,
        }),
        ExprKind::IsType { expr, type_name } => {
            names.contains(type_name) || java_expr_references_any_name(expr, names)
        }
        ExprKind::DefaultOf(type_name) => names.contains(type_name),
        ExprKind::NullCoalesce { left, right } => {
            java_expr_references_any_name(left, names)
                || java_expr_references_any_name(right, names)
        }
        ExprKind::SuperCall { args, .. } => args
            .iter()
            .any(|arg| java_expr_references_any_name(&arg.value, names)),
        ExprKind::Comprehension {
            element,
            generators,
            ..
        } => {
            java_expr_references_any_name(element, names)
                || generators.iter().any(|generator| {
                    java_expr_references_any_name(&generator.iter, names)
                        || generator
                            .conditions
                            .iter()
                            .any(|expr| java_expr_references_any_name(expr, names))
                })
        }
        ExprKind::Slice { lower, upper, step } => {
            lower
                .as_ref()
                .is_some_and(|expr| java_expr_references_any_name(expr, names))
                || upper
                    .as_ref()
                    .is_some_and(|expr| java_expr_references_any_name(expr, names))
                || step
                    .as_ref()
                    .is_some_and(|expr| java_expr_references_any_name(expr, names))
        }
        ExprKind::Walrus { target, value } => {
            java_expr_references_any_name(target, names)
                || java_expr_references_any_name(value, names)
        }
        ExprKind::ClassExpr {
            parent, members, ..
        } => {
            parent
                .as_ref()
                .is_some_and(|expr| java_expr_references_any_name(expr, names))
                || members
                    .iter()
                    .any(|member| java_class_member_references_any_name(member, names))
        }
        ExprKind::FunctionExpr(stmt) => java_stmt_references_any_name(stmt, names),
        ExprKind::Range { start, end, .. } => {
            java_expr_references_any_name(start, names) || java_expr_references_any_name(end, names)
        }
        ExprKind::StaticAccess { class, member } => {
            java_expr_references_any_name(class, names)
                || java_expr_references_any_name(member, names)
        }
        _ => false,
    }
}

// ════════════════════════════════════════════════════════════════════════════
// Span / helpers
// ════════════════════════════════════════════════════════════════════════════

fn to_span(pair: &Pair<Rule>) -> Span {
    let (line, col) = pair.line_col();
    Span {
        start_line: line as u32,
        start_col: col as u32,
        end_line: line as u32,
        end_col: col as u32,
    }
}

fn is_kw(rule: Rule) -> bool {
    matches!(
        rule,
        Rule::final_kw
            | Rule::static_kw
            | Rule::public_kw
            | Rule::private_kw
            | Rule::protected_kw
            | Rule::abstract_kw
            | Rule::synchronized_kw
            | Rule::native_kw
            | Rule::transient_kw
            | Rule::volatile_kw
            | Rule::strictfp_kw
            | Rule::default_kw
            | Rule::sealed_kw
            | Rule::non_sealed_kw
            | Rule::var_kw
    )
}

// ════════════════════════════════════════════════════════════════════════════
// Imports
// ════════════════════════════════════════════════════════════════════════════

/// JLS §7.5.3/§7.5.4: static imports bind a type's static member names
/// into the compilation unit. Walker-resolved (frontend-first): a bare
/// call `max(2, 9)` under `import static java.lang.Math.max;` rewrites
/// to `Math.max(2, 9)`, which the existing builtin/known-type dispatch
/// already compiles. On-demand (`Math.*`) rewrites bare calls whose
/// names are neither declared in the unit, single-imported, nor bare
/// profile builtins; with several on-demand imports the first one wins
/// (real javac rejects genuinely ambiguous uses — not modeled).
fn rewrite_java_static_imports(
    body: &mut [Statement],
    singles: &HashMap<String, String>,
    on_demand: &[String],
) {
    if singles.is_empty() && on_demand.is_empty() {
        return;
    }
    let mut declared = HashSet::new();
    collect_java_declared_callables(body, &mut declared);
    rewrite_static_import_stmts(body, singles, on_demand, &declared);
}

/// Method/function names declared anywhere in the unit — those shadow
/// on-demand static imports (JLS §6.4.1).
fn collect_java_declared_callables(stmts: &[Statement], out: &mut HashSet<String>) {
    for stmt in stmts {
        match &stmt.kind {
            StmtKind::FunctionDecl { name, body, .. } => {
                out.insert(name.clone());
                collect_java_declared_callables(body, out);
            }
            StmtKind::ClassDecl { members, .. } => {
                for member in members {
                    match member {
                        ClassMember::Method(method) => {
                            if let StmtKind::FunctionDecl { name, body, .. } = &method.kind {
                                out.insert(name.clone());
                                collect_java_declared_callables(body, out);
                            }
                        }
                        ClassMember::NestedType(nested) => {
                            collect_java_declared_callables(std::slice::from_ref(nested), out);
                        }
                        _ => {}
                    }
                }
            }
            StmtKind::Block(stmts) => collect_java_declared_callables(stmts, out),
            _ => {}
        }
    }
}

fn rewrite_static_import_stmts(
    stmts: &mut [Statement],
    singles: &HashMap<String, String>,
    on_demand: &[String],
    declared: &HashSet<String>,
) {
    for stmt in stmts {
        rewrite_static_import_stmt(stmt, singles, on_demand, declared);
    }
}

fn rewrite_static_import_stmt(
    stmt: &mut Statement,
    singles: &HashMap<String, String>,
    on_demand: &[String],
    declared: &HashSet<String>,
) {
    let e = |expr: &mut Expression| rewrite_static_import_expr(expr, singles, on_demand, declared);
    let b =
        |body: &mut [Statement]| rewrite_static_import_stmts(body, singles, on_demand, declared);
    match &mut stmt.kind {
        StmtKind::Expr(expr) => e(expr),
        StmtKind::Block(stmts) => b(stmts),
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    e(init);
                }
            }
        }
        StmtKind::FunctionDecl { body, .. } => b(body),
        StmtKind::ClassDecl { members, .. } => {
            for member in members {
                match member {
                    ClassMember::Method(method) | ClassMember::NestedType(method) => {
                        rewrite_static_import_stmt(method, singles, on_demand, declared);
                    }
                    ClassMember::Constructor { body, .. } => b(body),
                    ClassMember::Field {
                        init: Some(init), ..
                    } => e(init),
                    _ => {}
                }
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            e(cond);
            b(then_body);
            for (elif_cond, elif_body) in elifs {
                e(elif_cond);
                b(elif_body);
            }
            if let Some(else_body) = else_body {
                b(else_body);
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                rewrite_static_import_stmt(init, singles, on_demand, declared);
            }
            if let Some(cond) = cond {
                e(cond);
            }
            if let Some(update) = update {
                e(update);
            }
            b(body);
        }
        StmtKind::ForIn { iter, body, .. } => {
            e(iter);
            b(body);
        }
        StmtKind::While { cond, body, .. } => {
            e(cond);
            b(body);
        }
        StmtKind::DoWhile { body, cond, .. } => {
            b(body);
            e(cond);
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            e(expr);
            // Case conditions are constant expressions in Java — no
            // static-import call rewriting needed there.
            for case in cases {
                b(&mut case.body);
            }
            if let Some(default) = default {
                b(default);
            }
        }
        StmtKind::Try {
            body,
            catches,
            finally,
            ..
        } => {
            b(body);
            for catch in catches {
                b(&mut catch.body);
            }
            if let Some(finally) = finally {
                b(finally);
            }
        }
        StmtKind::Return(Some(expr)) => e(expr),
        StmtKind::Throw {
            expr: Some(expr), ..
        } => e(expr),
        StmtKind::Assign { targets, value } => {
            for target in targets {
                e(target);
            }
            e(value);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            e(target);
            e(value);
        }
        _ => {}
    }
}

fn rewrite_static_import_expr(
    expr: &mut Expression,
    singles: &HashMap<String, String>,
    on_demand: &[String],
    declared: &HashSet<String>,
) {
    // Rewrite the bare-call head FIRST, then recurse into children.
    if let ExprKind::Call { callee, .. } = &mut expr.kind {
        if let ExprKind::Ident(name) = &callee.kind {
            let receiver = singles.get(name.as_str()).cloned().or_else(|| {
                (!declared.contains(name.as_str())
                    && !name.starts_with("__")
                    && !matches!(name.as_str(), "println" | "print" | "printf"))
                .then(|| on_demand.first().cloned())
                .flatten()
            });
            if let Some(ty) = receiver {
                let field = name.clone();
                let span = callee.span.clone();
                **callee = Expression::with_span(
                    ExprKind::Member {
                        object: Box::new(Expression::new(ExprKind::Ident(ty))),
                        field,
                        null_safe: false,
                    },
                    span,
                );
            }
        }
    }
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            rewrite_static_import_expr(callee, singles, on_demand, declared);
            for arg in &mut *args {
                rewrite_static_import_expr(&mut arg.value, singles, on_demand, declared);
            }
        }
        ExprKind::New { args, .. } | ExprKind::SuperCall { args, .. } => {
            for arg in args {
                rewrite_static_import_expr(&mut arg.value, singles, on_demand, declared);
            }
        }
        ExprKind::Member { object, .. } => {
            rewrite_static_import_expr(object, singles, on_demand, declared);
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_static_import_expr(object, singles, on_demand, declared);
            rewrite_static_import_expr(index, singles, on_demand, declared);
        }
        ExprKind::Binary { left, right, .. } | ExprKind::NullCoalesce { left, right } => {
            rewrite_static_import_expr(left, singles, on_demand, declared);
            rewrite_static_import_expr(right, singles, on_demand, declared);
        }
        ExprKind::Unary { expr: inner, .. }
        | ExprKind::Spread(inner)
        | ExprKind::Await(inner)
        | ExprKind::TypeOf(inner) => {
            rewrite_static_import_expr(inner, singles, on_demand, declared);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_static_import_expr(cond, singles, on_demand, declared);
            rewrite_static_import_expr(then, singles, on_demand, declared);
            rewrite_static_import_expr(else_, singles, on_demand, declared);
        }
        ExprKind::Assign { target, value } => {
            rewrite_static_import_expr(target, singles, on_demand, declared);
            rewrite_static_import_expr(value, singles, on_demand, declared);
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(inner) => {
                rewrite_static_import_expr(inner, singles, on_demand, declared)
            }
            LambdaBody::Block(stmts) => {
                rewrite_static_import_stmts(stmts, singles, on_demand, declared)
            }
        },
        ExprKind::Array(elements) => {
            for element in elements {
                rewrite_static_import_expr(&mut element.value, singles, on_demand, declared);
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) => {
            for item in items {
                rewrite_static_import_expr(item, singles, on_demand, declared);
            }
        }
        _ => {}
    }
}

fn walk_import(pair: &Pair<Rule>) -> Option<Import> {
    let span = to_span(pair);
    let src = pair.as_str();
    let text = src
        .trim_start_matches("import")
        .trim_start_matches(" static")
        .trim()
        .trim_end_matches(';')
        .trim();
    if text.is_empty() {
        return None;
    }
    // `import java.util.*;` — the package is the namespace. Recorded as a
    // Wildcard over the dotted package path (namespaceplan.md ambient
    // shape); non-host wildcard paths are inert at link time, so this is
    // data, not behavior — name resolution stays with [builtins] /
    // known_types until ambient-from-imports plumbing lands.
    if let Some(package) = text.strip_suffix(".*") {
        return Some(Import {
            kind: ImportKind::Wildcard {
                path: package.to_string(),
                alias: None,
            },
            span,
        });
    }
    // `import java.util.HashMap;` — bind the simple name to its
    // fully-qualified dotted path via `Simple{alias}` (the same
    // `source_type_aliases` map instanceof/static-access resolve
    // through). The previous shape here synthesized a Named import
    // against a nonexistent `java:java/util/…` module specifier, which
    // poisoned `host_import_bindings` for every imported simple name.
    let name = text.rsplit('.').next().unwrap_or(text).to_string();
    Some(Import {
        kind: ImportKind::Simple {
            path: text.to_string(),
            alias: Some(name),
        },
        span,
    })
}

// ════════════════════════════════════════════════════════════════════════════
// Modifiers
// ════════════════════════════════════════════════════════════════════════════

#[derive(Clone, Copy)]
struct ParsedModifiers {
    visibility: Visibility,
    is_static: bool,
    is_abstract: bool,
    is_final: bool,
}

impl Default for ParsedModifiers {
    fn default() -> Self {
        Self {
            visibility: Visibility::Public,
            is_static: false,
            is_abstract: false,
            is_final: false,
        }
    }
}

fn parse_modifiers(pair: &Pair<Rule>) -> ParsedModifiers {
    let mut out = ParsedModifiers::default();
    if pair.as_rule() != Rule::modifiers {
        return out;
    }
    for m in pair.clone().into_inner() {
        if m.as_rule() == Rule::modifier {
            let inner = m.into_inner().next();
            match inner.as_ref().map(|p| p.as_rule()) {
                Some(Rule::private_kw) => out.visibility = Visibility::Private,
                Some(Rule::protected_kw) => out.visibility = Visibility::Protected,
                Some(Rule::public_kw) => out.visibility = Visibility::Public,
                Some(Rule::static_kw) => out.is_static = true,
                Some(Rule::abstract_kw) => out.is_abstract = true,
                Some(Rule::final_kw) => out.is_final = true,
                _ => {}
            }
        }
    }
    out
}

fn into_modifiers(pm: ParsedModifiers) -> Modifiers {
    Modifiers {
        visibility: pm.visibility,
        is_static: pm.is_static,
        is_abstract: pm.is_abstract,
        is_readonly: pm.is_final,
        ..Modifiers::default()
    }
}

fn into_class_modifiers(pm: ParsedModifiers) -> ClassModifiers {
    ClassModifiers {
        visibility: pm.visibility,
        is_abstract: pm.is_abstract,
        is_static: pm.is_static,
        ..ClassModifiers::default()
    }
}

fn java_visibility_modifier_bits(visibility: Visibility) -> i64 {
    match visibility {
        Visibility::Public => JAVA_MOD_PUBLIC,
        Visibility::Private => JAVA_MOD_PRIVATE,
        Visibility::Protected => JAVA_MOD_PROTECTED,
        Visibility::Internal => 0,
    }
}

fn java_parsed_modifier_bits(pm: ParsedModifiers) -> i64 {
    let mut bits = java_visibility_modifier_bits(pm.visibility);
    if pm.is_static {
        bits |= JAVA_MOD_STATIC;
    }
    if pm.is_final {
        bits |= JAVA_MOD_FINAL;
    }
    if pm.is_abstract {
        bits |= JAVA_MOD_ABSTRACT;
    }
    bits
}

fn java_member_modifier_bits(modifiers: &Modifiers) -> i64 {
    let mut bits = java_visibility_modifier_bits(modifiers.visibility);
    if modifiers.is_static {
        bits |= JAVA_MOD_STATIC;
    }
    if modifiers.is_readonly {
        bits |= JAVA_MOD_FINAL;
    }
    if modifiers.is_abstract {
        bits |= JAVA_MOD_ABSTRACT;
    }
    bits
}

fn java_class_modifier_bits(modifiers: &ClassModifiers) -> i64 {
    let mut bits = java_visibility_modifier_bits(modifiers.visibility);
    if modifiers.is_static {
        bits |= JAVA_MOD_STATIC;
    }
    if modifiers.is_sealed {
        bits |= JAVA_MOD_FINAL;
    }
    if modifiers.is_abstract {
        bits |= JAVA_MOD_ABSTRACT;
    }
    bits
}

fn java_modifier_static_predicate(method: &str, value: &Expression) -> Option<Expression> {
    let mask = match method {
        "isPublic" => JAVA_MOD_PUBLIC,
        "isPrivate" => JAVA_MOD_PRIVATE,
        "isProtected" => JAVA_MOD_PROTECTED,
        "isStatic" => JAVA_MOD_STATIC,
        "isFinal" => JAVA_MOD_FINAL,
        "isAbstract" => JAVA_MOD_ABSTRACT,
        _ => return None,
    };
    let ExprKind::Lit(Literal::Int(bits)) = &value.kind else {
        return None;
    };
    Some(Expression::bool((*bits & mask) != 0))
}

// ════════════════════════════════════════════════════════════════════════════
// Class
// ════════════════════════════════════════════════════════════════════════════

fn consume_java_type_params(pair: Pair<Rule>) {
    let _ = common_generics::parse_generic_params_hint(pair.as_str());
}

fn walk_class(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut parents: Vec<String> = Vec::new();
    let mut interfaces: Vec<String> = Vec::new();
    let mut members: Vec<ClassMember> = Vec::new();
    let mut class_modifiers = ClassModifiers::default();
    let mut class_modifier_bits = JAVA_MOD_PUBLIC;

    let mut inner = pair.into_inner().peekable();

    // modifiers
    if inner.peek().map(|p| p.as_rule()) == Some(Rule::modifiers) {
        let mp = inner.next().unwrap();
        let parsed = parse_modifiers(&mp);
        class_modifier_bits = java_parsed_modifier_bits(parsed);
        class_modifiers = into_class_modifiers(parsed);
    }
    // "class" keyword matched as ident_name in grammar — next ident is the name
    if let Some(n) = inner.next() {
        name = n.as_str().to_string();
    }

    for p in inner {
        match p.as_rule() {
            Rule::type_params => consume_java_type_params(p),
            Rule::type_ref => {
                // extends clause: first type_ref
                if parents.is_empty() {
                    parents.push(extract_ref_name(&p));
                }
            }
            Rule::type_ref_list => {
                for tr in p.into_inner() {
                    if tr.as_rule() == Rule::type_ref {
                        interfaces.push(extract_ref_name(&tr));
                    }
                }
            }
            Rule::class_body => {
                members = walk_class_body_with_owner(p, Some(&name))?;
            }
            _ => {}
        }
    }

    let extends_thread = parents
        .iter()
        .any(|parent| java_type_simple_name(parent) == "Thread");
    if extends_thread {
        parents.retain(|parent| java_type_simple_name(parent) != "Thread");
        JAVA_THREAD_TYPES.with(|types| {
            types.borrow_mut().insert(name.clone());
        });
        inject_java_thread_stamps(&mut members);
    } else if !parents.is_empty() {
        inject_implicit_super(&mut members);
    }

    // Custom exception classes (`class X extends RuntimeException`): the
    // parent is a built-in with no real class behind it — stamp the
    // canonical exception shape in every ctor and turn `super(msg[,cause])`
    // into message/cause stores.
    if let Some(parent) = parents.first() {
        let parent_simple = java_type_simple_name(parent).to_string();
        if let Some(chain) = java_exception_supertypes(&parent_simple) {
            inject_java_exception_stamps(&name, &chain, &mut members);
        }
    }

    JAVA_REFLECTION_CLASSES.with(|classes| {
        classes.borrow_mut().insert(
            name.clone(),
            java_reflection_meta(
                &name,
                parents.first().cloned(),
                interfaces.clone(),
                &members,
                false,
                false,
                class_modifier_bits,
            ),
        );
    });

    Ok(StmtKind::ClassDecl {
        name,
        parents,
        interfaces,
        members,
        modifiers: class_modifiers,
        decorators: vec![],
    })
}

fn inject_java_exception_stamps(
    class_name: &str,
    chain: &[String],
    members: &mut Vec<ClassMember>,
) {
    let this_field = |f: &str| {
        Expression::new(ExprKind::Member {
            object: Box::new(Expression::new(ExprKind::This)),
            field: f.to_string(),
            null_safe: false,
        })
    };
    let assign_stmt = |target: Expression, value: Expression| {
        Statement::new(StmtKind::Assign {
            targets: vec![target],
            value,
        })
    };
    let mut types_elems = vec![ArrayElement {
        key: None,
        value: Expression::string(class_name),
        spread: false,
        by_ref: false,
    }];
    types_elems.extend(chain.iter().map(|t| ArrayElement {
        key: None,
        value: Expression::string(t),
        spread: false,
        by_ref: false,
    }));

    let mut has_ctor = false;
    for member in members.iter_mut() {
        if let ClassMember::Constructor {
            body, base_args, ..
        } = member
        {
            has_ctor = true;
            let mut prelude: Vec<Statement> = vec![
                assign_stmt(
                    this_field("__exception_type"),
                    Expression::string(class_name),
                ),
                assign_stmt(this_field("name"), Expression::string(class_name)),
                assign_stmt(
                    this_field("__types"),
                    Expression::new(ExprKind::Array(types_elems.clone())),
                ),
                assign_stmt(this_field("message"), Expression::null()),
            ];
            if let Some(args) = base_args.take() {
                if let Some(msg) = args.first() {
                    prelude.push(assign_stmt(this_field("message"), msg.clone()));
                }
                if let Some(cause) = args.get(1) {
                    prelude.push(assign_stmt(this_field("cause"), cause.clone()));
                }
            }
            *base_args = None;
            prelude.append(body);
            *body = prelude;
        }
    }
    if !has_ctor {
        members.push(ClassMember::Constructor {
            name: None,
            params: vec![],
            body: vec![
                assign_stmt(
                    this_field("__exception_type"),
                    Expression::string(class_name),
                ),
                assign_stmt(this_field("name"), Expression::string(class_name)),
                assign_stmt(
                    this_field("__types"),
                    Expression::new(ExprKind::Array(types_elems)),
                ),
                assign_stmt(this_field("message"), Expression::null()),
            ],
            base_args: None,
            initializer_target: ConstructorInitializerTarget::Base,
            visibility: ParsedModifiers::default().visibility,
        });
    }
}

fn walk_class_body(pair: Pair<Rule>) -> Result<Vec<ClassMember>, String> {
    walk_class_body_with_owner(pair, None)
}

fn walk_class_body_with_owner(
    pair: Pair<Rule>,
    owner: Option<&str>,
) -> Result<Vec<ClassMember>, String> {
    let mut members = Vec::new();
    if let Some(owner) = owner {
        JAVA_CURRENT_CLASS_STACK.with(|stack| stack.borrow_mut().push(owner.to_string()));
    }
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::constructor_declaration => members.push(walk_constructor(p)?),
            Rule::method_declaration | Rule::default_method_declaration => {
                members.push(walk_method(p)?)
            }
            Rule::field_declaration => {
                for m in walk_field(p)? {
                    if let (
                        Some(owner),
                        ClassMember::Field {
                            name,
                            type_hint,
                            modifiers,
                            ..
                        },
                    ) = (owner, &m)
                    {
                        if modifiers.is_static {
                            let key = format!("{owner}.{name}");
                            JAVA_STATIC_FIELD_VARS.with(|vars| {
                                vars.borrow_mut().insert(key.clone());
                            });
                            if let Some(type_hint) = type_hint {
                                JAVA_STATIC_FIELD_TYPES.with(|types| {
                                    types.borrow_mut().insert(key, type_hint.clone());
                                });
                            }
                        }
                    }
                    members.push(m);
                }
            }
            Rule::class_declaration => {
                members.push(ClassMember::NestedType(Box::new(Statement::new(
                    walk_class(p)?,
                ))));
            }
            Rule::interface_declaration => {
                members.push(ClassMember::NestedType(Box::new(Statement::new(
                    walk_interface(p)?,
                ))));
            }
            Rule::enum_declaration => {
                members.push(ClassMember::NestedType(Box::new(Statement::new(
                    walk_enum_decl(p)?,
                ))));
            }
            Rule::record_declaration => {
                members.push(ClassMember::NestedType(Box::new(Statement::new(
                    walk_record(p)?,
                ))));
            }
            Rule::static_initializer | Rule::instance_initializer => {
                let is_static = p.as_rule() == Rule::static_initializer;
                let body: Vec<Statement> = p
                    .into_inner()
                    .filter_map(|b| {
                        if b.as_rule() == Rule::block_statement {
                            walk_block(b).ok()
                        } else {
                            None
                        }
                    })
                    .flatten()
                    .collect();
                let mut modifiers = Modifiers::default();
                modifiers.is_static = is_static;
                members.push(ClassMember::Method(Box::new(Statement::new(
                    StmtKind::FunctionDecl {
                        name: if is_static {
                            "__static_init_block__".to_string()
                        } else {
                            "__init_block__".to_string()
                        },
                        params: vec![],
                        return_type: None,
                        body,
                        modifiers,
                        handles: vec![],
                        is_async: false,
                        is_generator: false,
                        is_sub: false,
                    },
                ))));
            }
            Rule::annotation => {}
            _ => {}
        }
    }
    if owner.is_some() {
        JAVA_CURRENT_CLASS_STACK.with(|stack| {
            stack.borrow_mut().pop();
        });
    }
    Ok(members)
}

fn walk_constructor(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut inner = pair.into_inner().peekable();

    let pm = if inner.peek().map(|p| p.as_rule()) == Some(Rule::modifiers) {
        parse_modifiers(&inner.next().unwrap())
    } else {
        ParsedModifiers::default()
    };
    let visibility = pm.visibility;

    // Java generic constructors write type params before the constructor name:
    // `<T> Box(T value)`. They normalize through the shared generics primitive,
    // then erase at runtime.
    if inner.peek().map(|p| p.as_rule()) == Some(Rule::type_params) {
        consume_java_type_params(inner.next().unwrap());
    }

    // constructor name — same as class, skip
    if inner.peek().map(|p| p.as_rule()) == Some(Rule::ident_name) {
        inner.next();
    }
    if inner.peek().map(|p| p.as_rule()) == Some(Rule::type_params) {
        consume_java_type_params(inner.next().unwrap());
    }

    let mut params: Vec<Param> = Vec::new();
    let mut body: Vec<Statement> = Vec::new();
    let mut base_args: Option<Vec<Expression>> = None;
    let mut initializer_target = ConstructorInitializerTarget::Base;

    for p in inner {
        match p.as_rule() {
            Rule::param_list => params = walk_params(p)?,
            Rule::throws_clause => {}
            Rule::function_body_block => {
                body = walk_block(p)?;
                // Extract super(...) or this(...) call from top of body
                extract_base_call_from_body(&mut body, &mut base_args, &mut initializer_target);
            }
            _ => {}
        }
    }

    Ok(ClassMember::Constructor {
        name: None,
        params,
        body,
        base_args,
        initializer_target,
        visibility,
    })
}

fn walk_method(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut inner = pair.into_inner().peekable();

    let pm = if inner.peek().map(|p| p.as_rule()) == Some(Rule::modifiers) {
        parse_modifiers(&inner.next().unwrap())
    } else {
        ParsedModifiers::default()
    };
    let modifiers = into_modifiers(pm);

    // Java generic methods write type params before the return type:
    // `static <T> T identity(T x)`. They normalize through the shared generics
    // primitive, then erase at runtime.
    if inner.peek().map(|p| p.as_rule()) == Some(Rule::type_params) {
        consume_java_type_params(inner.next().unwrap());
    }

    // Return type (type_ref)
    let return_type = if inner.peek().map(|p| p.as_rule()) == Some(Rule::type_ref) {
        let tr = inner.next().unwrap();
        Some(extract_ref_name(&tr))
    } else {
        None
    };

    let name = inner
        .next()
        .ok_or("method: missing name")?
        .as_str()
        .to_string();

    if inner.peek().map(|p| p.as_rule()) == Some(Rule::type_params) {
        consume_java_type_params(inner.next().unwrap());
    }

    let mut params: Vec<Param> = Vec::new();
    let mut body: Vec<Statement> = Vec::new();

    for p in inner {
        match p.as_rule() {
            Rule::param_list => params = walk_params(p)?,
            Rule::dim_suffix => {}
            Rule::throws_clause => {}
            Rule::function_body => {
                for fb in p.into_inner() {
                    if fb.as_rule() == Rule::function_body_block {
                        body = walk_block(fb)?;
                    }
                }
            }
            _ => {}
        }
    }

    Ok(ClassMember::Method(Box::new(Statement::new(
        StmtKind::FunctionDecl {
            name,
            params,
            return_type,
            body,
            modifiers,
            handles: vec![],
            is_async: false,
            is_generator: false,
            is_sub: false,
        },
    ))))
}

fn walk_field(pair: Pair<Rule>) -> Result<Vec<ClassMember>, String> {
    let mut inner = pair.into_inner().peekable();

    let pm = if inner.peek().map(|p| p.as_rule()) == Some(Rule::modifiers) {
        parse_modifiers(&inner.next().unwrap())
    } else {
        ParsedModifiers::default()
    };
    let modifiers = into_modifiers(pm);

    // type_ref
    let type_hint = if inner.peek().map(|p| p.as_rule()) == Some(Rule::type_ref) {
        Some(extract_ref_name(&inner.next().unwrap()))
    } else {
        None
    };

    let mut fields = Vec::new();
    for p in inner {
        if p.as_rule() == Rule::var_declarator {
            let mut di = p.into_inner().peekable();
            let name = di.next().ok_or("field: missing name")?.as_str().to_string();
            // skip dim_suffix(s)
            while di.peek().map(|x| x.as_rule()) == Some(Rule::dim_suffix) {
                di.next();
            }
            let init = if di.peek().map(|x| x.as_rule()) == Some(Rule::initializer) {
                Some(walk_initializer(di.next().unwrap())?)
            } else {
                type_hint.as_deref().and_then(default_expr_for_java_type)
            };
            fields.push(ClassMember::Field {
                name,
                type_hint: type_hint.clone(),
                init,
                modifiers: modifiers.clone(),
                with_events: false,
                array_bounds: None,
            });
        }
    }
    Ok(fields)
}

fn java_single_abstract_interface_method(members: &[ClassMember]) -> Option<String> {
    let mut abstract_methods = members.iter().filter_map(|member| {
        let ClassMember::Method(func) = member else {
            return None;
        };
        let StmtKind::FunctionDecl {
            name,
            body,
            modifiers,
            ..
        } = &func.kind
        else {
            return None;
        };
        (!modifiers.is_static && body.is_empty()).then(|| name.clone())
    });
    let method = abstract_methods.next()?;
    abstract_methods.next().is_none().then_some(method)
}

// ════════════════════════════════════════════════════════════════════════════
// Interface
// ════════════════════════════════════════════════════════════════════════════

fn walk_interface(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner().peekable();
    let pm = if inner.peek().map(|p| p.as_rule()) == Some(Rule::modifiers) {
        parse_modifiers(&inner.next().unwrap())
    } else {
        ParsedModifiers::default()
    };
    let name = inner
        .next()
        .ok_or("interface: missing name")?
        .as_str()
        .to_string();
    JAVA_INTERFACE_NAMES.with(|names| {
        names.borrow_mut().insert(name.clone());
    });
    if inner.peek().map(|p| p.as_rule()) == Some(Rule::type_params) {
        consume_java_type_params(inner.next().unwrap());
    }

    let mut parents: Vec<String> = Vec::new();
    let mut members: Vec<ClassMember> = Vec::new();

    for p in inner {
        match p.as_rule() {
            Rule::type_ref_list => {
                for tr in p.into_inner() {
                    if tr.as_rule() == Rule::type_ref {
                        parents.push(extract_ref_name(&tr));
                    }
                }
            }
            Rule::interface_body => {
                for m in p.into_inner() {
                    match m.as_rule() {
                        Rule::method_declaration | Rule::default_method_declaration => {
                            members.push(walk_method(m)?);
                        }
                        Rule::field_declaration => {
                            for f in walk_field(m)? {
                                members.push(f);
                            }
                        }
                        _ => {}
                    }
                }
            }
            _ => {}
        }
    }

    JAVA_REFLECTION_CLASSES.with(|classes| {
        let modifier_bits = java_parsed_modifier_bits(pm) | JAVA_MOD_ABSTRACT;
        classes.borrow_mut().insert(
            name.clone(),
            java_reflection_meta(
                &name,
                parents.first().cloned(),
                vec![],
                &members,
                true,
                false,
                modifier_bits,
            ),
        );
    });
    if let Some(method_name) = java_single_abstract_interface_method(&members) {
        JAVA_FUNCTIONAL_INTERFACE_METHODS.with(|methods| {
            methods.borrow_mut().insert(name.clone(), method_name);
        });
    }

    Ok(StmtKind::ClassDecl {
        name,
        parents,
        interfaces: vec![],
        members,
        modifiers: into_class_modifiers(pm),
        decorators: vec![],
    })
}

// ════════════════════════════════════════════════════════════════════════════
// Enum
// ════════════════════════════════════════════════════════════════════════════

fn walk_enum_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner().peekable();
    let pm = if inner.peek().map(|p| p.as_rule()) == Some(Rule::modifiers) {
        parse_modifiers(&inner.next().unwrap())
    } else {
        ParsedModifiers::default()
    };
    let name = inner
        .next()
        .ok_or("enum: missing name")?
        .as_str()
        .to_string();

    let mut enum_members: Vec<EnumMember> = Vec::new();
    let mut member_ctor_args: Vec<Vec<Argument>> = Vec::new();
    let mut member_overrides: Vec<(String, Vec<ClassMember>)> = Vec::new();
    let mut interfaces: Vec<String> = Vec::new();
    let mut body_members: Vec<ClassMember> = Vec::new();

    for p in inner {
        match p.as_rule() {
            Rule::type_ref_list => {
                for tr in p.into_inner() {
                    if tr.as_rule() == Rule::type_ref {
                        interfaces.push(extract_ref_name(&tr));
                    }
                }
            }
            Rule::enum_values => {
                for ev in p.into_inner() {
                    if ev.as_rule() == Rule::enum_value {
                        let mut val_name = String::new();
                        let mut args: Vec<Argument> = Vec::new();
                        let mut overrides: Vec<ClassMember> = Vec::new();
                        for x in ev.into_inner() {
                            match x.as_rule() {
                                Rule::ident_name => val_name = x.as_str().to_string(),
                                Rule::argument_list => args = walk_arguments(x)?,
                                Rule::class_body => overrides = walk_class_body(x)?,
                                _ => {}
                            }
                        }
                        if !val_name.is_empty() {
                            enum_members.push(EnumMember {
                                name: val_name.clone(),
                                value: None,
                                constructor_args: args.iter().map(|a| a.value.clone()).collect(),
                            });
                            member_ctor_args.push(args);
                            if !overrides.is_empty() {
                                member_overrides.push((val_name, overrides));
                            }
                        }
                    }
                }
            }
            // Members after the `;` — `class_member` is a silent rule, so the
            // declarations appear directly here.
            Rule::constructor_declaration => body_members.push(walk_constructor(p)?),
            Rule::method_declaration | Rule::default_method_declaration => {
                body_members.push(walk_method(p)?)
            }
            Rule::field_declaration => body_members.extend(walk_field(p)?),
            Rule::class_declaration => {
                body_members.push(ClassMember::NestedType(Box::new(Statement::new(
                    walk_class(p)?,
                ))));
            }
            _ => {}
        }
    }

    apply_java_enum_constant_method_overrides(&name, &mut body_members, &member_overrides);

    JAVA_ENUM_VALUES.with(|values| {
        values.borrow_mut().insert(
            name.clone(),
            enum_members
                .iter()
                .map(|member| member.name.clone())
                .collect(),
        );
    });

    // JLS §8.9: each constant is an instance of the enum class. Constants
    // become static instances (`Season.SPRING = new Season("SPRING", 0, …)`);
    // name()/ordinal()/toString()/values()/valueOf() are synthesized unless
    // the body declares them.
    let simple_param = |n: &str| Param {
        name: n.to_string(),
        type_hint: None,
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    };
    let this_field = |f: &str| {
        Expression::new(ExprKind::Member {
            object: Box::new(Expression::new(ExprKind::This)),
            field: f.to_string(),
            null_safe: false,
        })
    };
    let stamp = |field: &str, param: &str| {
        Statement::new(StmtKind::Assign {
            targets: vec![this_field(field)],
            value: Expression::ident(param),
        })
    };

    let mut has_ctor = false;
    for member in body_members.iter_mut() {
        if let ClassMember::Constructor { params, body, .. } = member {
            has_ctor = true;
            params.insert(0, simple_param("__ordinal"));
            params.insert(0, simple_param("__name"));
            body.insert(0, stamp("__ordinal", "__ordinal"));
            body.insert(0, stamp("__name", "__name"));
        }
    }
    if !has_ctor {
        body_members.push(ClassMember::Constructor {
            name: None,
            params: vec![simple_param("__name"), simple_param("__ordinal")],
            body: vec![stamp("__name", "__name"), stamp("__ordinal", "__ordinal")],
            base_args: None,
            initializer_target: ConstructorInitializerTarget::Base,
            visibility: ParsedModifiers::default().visibility,
        });
    }

    let user_method_names: Vec<String> = body_members
        .iter()
        .filter_map(|m| match m {
            ClassMember::Method(s) => match &s.kind {
                StmtKind::FunctionDecl { name, .. } => Some(name.clone()),
                _ => None,
            },
            _ => None,
        })
        .collect();
    let make_method = |mname: &str, params: Vec<Param>, body: Vec<Statement>, is_static: bool| {
        let mut modifiers = Modifiers::default();
        modifiers.is_static = is_static;
        ClassMember::Method(Box::new(Statement::new(StmtKind::FunctionDecl {
            name: mname.to_string(),
            params,
            return_type: None,
            body,
            modifiers,
            handles: vec![],
            is_async: false,
            is_generator: false,
            is_sub: false,
        })))
    };
    for (mname, field) in [
        ("name", "__name"),
        ("ordinal", "__ordinal"),
        ("toString", "__name"),
    ] {
        if !user_method_names.iter().any(|n| n == mname) {
            body_members.push(make_method(
                mname,
                vec![],
                vec![Statement::new(StmtKind::Return(Some(this_field(field))))],
                false,
            ));
        }
    }

    let member_access = |m: &str| {
        Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident(&name)),
            field: m.to_string(),
            null_safe: false,
        })
    };
    if !user_method_names.iter().any(|n| n == "values") {
        let values_array = Expression::new(ExprKind::Array(
            enum_members
                .iter()
                .map(|m| ArrayElement {
                    key: None,
                    value: member_access(&m.name),
                    spread: false,
                    by_ref: false,
                })
                .collect(),
        ));
        body_members.push(make_method(
            "values",
            vec![],
            vec![Statement::new(StmtKind::Return(Some(values_array)))],
            true,
        ));
    }
    if !user_method_names.iter().any(|n| n == "valueOf") {
        let mut body: Vec<Statement> = enum_members
            .iter()
            .map(|m| {
                Statement::new(StmtKind::If {
                    cond: java_binary(
                        BinOp::Eq,
                        Expression::ident("__s"),
                        Expression::string(&m.name),
                    ),
                    then_body: vec![Statement::new(StmtKind::Return(Some(member_access(
                        &m.name,
                    ))))],
                    elifs: vec![],
                    else_body: None,
                })
            })
            .collect();
        body.push(Statement::new(StmtKind::Throw {
            expr: Some(Expression::new(ExprKind::New {
                class: Box::new(Expression::ident("IllegalArgumentException")),
                args: vec![Argument::positional(java_binary(
                    BinOp::Add,
                    Expression::string(&format!("No enum constant {name}.")),
                    Expression::ident("__s"),
                ))],
            })),
            cause: None,
        }));
        // Not named `valueOf` — that name is intercepted by shared compiler
        // paths before the user-class static dispatch. The tostring post-pass
        // rewrites `EnumType.valueOf(x)` calls to this name.
        body_members.push(make_method(
            "__j_enum_value_of",
            vec![simple_param("__s")],
            body,
            true,
        ));
    }

    // Constants as static instance fields: `Mode.ON = new Mode("ON", 0, …)`.
    // Emitted as a plain ClassDecl — NOT StmtKind::EnumDecl — because the
    // shared EnumDecl path registers ordinal tables that constant-fold
    // `Mode.ON` member reads to F64 ordinals, breaking instance identity.
    for (i, m) in enum_members.iter().enumerate() {
        let mut args = vec![
            Argument::positional(Expression::string(&m.name)),
            Argument::positional(Expression::int(i as i64)),
        ];
        args.extend(member_ctor_args[i].clone());
        let mut modifiers = Modifiers::default();
        modifiers.is_static = true;
        body_members.push(ClassMember::Field {
            name: m.name.clone(),
            type_hint: Some(name.clone()),
            init: Some(Expression::new(ExprKind::New {
                class: Box::new(Expression::ident(&name)),
                args,
            })),
            modifiers,
            with_events: false,
            array_bounds: None,
        });
    }

    let enum_member_names: Vec<String> = enum_members
        .iter()
        .map(|member| member.name.clone())
        .collect();
    qualify_java_enum_intrinsics_in_members(&mut body_members, &name, &enum_member_names);

    // JLS §8.9: nested enum types are implicitly static — never capture an
    // outer instance (no `__java_outer` ctor param).
    let mut modifiers = into_class_modifiers(pm);
    modifiers.is_static = true;
    JAVA_REFLECTION_CLASSES.with(|classes| {
        let modifier_bits = java_parsed_modifier_bits(pm) | JAVA_MOD_STATIC | JAVA_MOD_FINAL;
        classes.borrow_mut().insert(
            name.clone(),
            java_reflection_meta(
                &name,
                Some("Enum".to_string()),
                interfaces.clone(),
                &body_members,
                false,
                true,
                modifier_bits,
            ),
        );
    });
    Ok(StmtKind::ClassDecl {
        name,
        parents: vec![],
        interfaces,
        members: body_members,
        modifiers,
        decorators: vec![],
    })
}

fn qualify_java_enum_intrinsics_in_members(
    members: &mut [ClassMember],
    enum_name: &str,
    enum_members: &[String],
) {
    for member in members {
        match member {
            ClassMember::Constructor { body, .. } => {
                qualify_java_enum_intrinsics_in_stmts(body, enum_name, enum_members);
            }
            ClassMember::Method(method) => {
                if let StmtKind::FunctionDecl { body, .. } = &mut method.kind {
                    qualify_java_enum_intrinsics_in_stmts(body, enum_name, enum_members);
                }
            }
            _ => {}
        }
    }
}

fn qualify_java_enum_intrinsics_in_stmts(
    stmts: &mut [Statement],
    enum_name: &str,
    enum_members: &[String],
) {
    for stmt in stmts {
        match &mut stmt.kind {
            StmtKind::Expr(expr) => {
                qualify_java_enum_intrinsics_in_expr(expr, enum_name, enum_members)
            }
            StmtKind::VarDecl { declarations, .. } => {
                for decl in declarations {
                    if let Some(init) = &mut decl.init {
                        qualify_java_enum_intrinsics_in_expr(init, enum_name, enum_members);
                    }
                }
            }
            StmtKind::Assign { targets, value } => {
                for target in targets {
                    qualify_java_enum_intrinsics_in_expr(target, enum_name, enum_members);
                }
                qualify_java_enum_intrinsics_in_expr(value, enum_name, enum_members);
            }
            StmtKind::Return(Some(expr))
            | StmtKind::Throw {
                expr: Some(expr), ..
            } => {
                qualify_java_enum_intrinsics_in_expr(expr, enum_name, enum_members);
            }
            StmtKind::If {
                cond,
                then_body,
                elifs,
                else_body,
            } => {
                qualify_java_enum_intrinsics_in_expr(cond, enum_name, enum_members);
                qualify_java_enum_intrinsics_in_stmts(then_body, enum_name, enum_members);
                for (elif_cond, elif_body) in elifs {
                    qualify_java_enum_intrinsics_in_expr(elif_cond, enum_name, enum_members);
                    qualify_java_enum_intrinsics_in_stmts(elif_body, enum_name, enum_members);
                }
                if let Some(else_body) = else_body {
                    qualify_java_enum_intrinsics_in_stmts(else_body, enum_name, enum_members);
                }
            }
            StmtKind::While { cond, body, .. } | StmtKind::DoWhile { cond, body, .. } => {
                qualify_java_enum_intrinsics_in_expr(cond, enum_name, enum_members);
                qualify_java_enum_intrinsics_in_stmts(body, enum_name, enum_members);
            }
            StmtKind::For {
                init,
                cond,
                update,
                body,
            } => {
                if let Some(init) = init {
                    qualify_java_enum_intrinsics_in_stmts(
                        std::slice::from_mut(init),
                        enum_name,
                        enum_members,
                    );
                }
                if let Some(cond) = cond {
                    qualify_java_enum_intrinsics_in_expr(cond, enum_name, enum_members);
                }
                if let Some(update) = update {
                    qualify_java_enum_intrinsics_in_expr(update, enum_name, enum_members);
                }
                qualify_java_enum_intrinsics_in_stmts(body, enum_name, enum_members);
            }
            StmtKind::ForIn {
                iter,
                body,
                else_body,
                ..
            } => {
                qualify_java_enum_intrinsics_in_expr(iter, enum_name, enum_members);
                qualify_java_enum_intrinsics_in_stmts(body, enum_name, enum_members);
                if let Some(else_body) = else_body {
                    qualify_java_enum_intrinsics_in_stmts(else_body, enum_name, enum_members);
                }
            }
            StmtKind::Block(body) => {
                qualify_java_enum_intrinsics_in_stmts(body, enum_name, enum_members)
            }
            StmtKind::Switch {
                expr,
                cases,
                default,
            } => {
                qualify_java_enum_intrinsics_in_expr(expr, enum_name, enum_members);
                for case in cases {
                    for condition in &mut case.conditions {
                        match condition {
                            CaseCondition::Value(value) => {
                                qualify_java_enum_intrinsics_in_expr(
                                    value,
                                    enum_name,
                                    enum_members,
                                );
                            }
                            CaseCondition::Range { from, to } => {
                                qualify_java_enum_intrinsics_in_expr(from, enum_name, enum_members);
                                qualify_java_enum_intrinsics_in_expr(to, enum_name, enum_members);
                            }
                            CaseCondition::Comparison { expr, .. } => {
                                qualify_java_enum_intrinsics_in_expr(expr, enum_name, enum_members);
                            }
                        }
                    }
                    qualify_java_enum_intrinsics_in_stmts(&mut case.body, enum_name, enum_members);
                }
                if let Some(default) = default {
                    qualify_java_enum_intrinsics_in_stmts(default, enum_name, enum_members);
                }
            }
            _ => {}
        }
    }
}

fn qualify_java_enum_intrinsics_in_expr(
    expr: &mut Expression,
    enum_name: &str,
    enum_members: &[String],
) {
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            if let ExprKind::Ident(name) = &callee.kind {
                if name == "values" && args.is_empty() {
                    *expr = Expression::new(ExprKind::Array(
                        enum_members
                            .iter()
                            .map(|member| ArrayElement {
                                key: None,
                                value: Expression::new(ExprKind::Member {
                                    object: Box::new(Expression::ident(enum_name)),
                                    field: member.clone(),
                                    null_safe: false,
                                }),
                                spread: false,
                                by_ref: false,
                            })
                            .collect(),
                    ));
                    return;
                }
                if name == "valueOf" {
                    *callee = Box::new(Expression::new(ExprKind::Member {
                        object: Box::new(Expression::ident(enum_name)),
                        field: name.clone(),
                        null_safe: false,
                    }));
                }
            }
            qualify_java_enum_intrinsics_in_expr(callee, enum_name, enum_members);
            for arg in args {
                qualify_java_enum_intrinsics_in_expr(&mut arg.value, enum_name, enum_members);
            }
        }
        ExprKind::Member { object, .. } => {
            qualify_java_enum_intrinsics_in_expr(object, enum_name, enum_members)
        }
        ExprKind::Index { object, index, .. } => {
            if let Some(member) =
                java_enum_values_constant_index_member(object, index, enum_members)
            {
                *expr = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Member {
                        object: Box::new(Expression::ident(enum_name)),
                        field: "__j_enum_value_of".to_string(),
                        null_safe: false,
                    })),
                    args: vec![Argument::positional(Expression::string(&member))],
                    optional: false,
                });
                return;
            }
            qualify_java_enum_intrinsics_in_expr(object, enum_name, enum_members);
            qualify_java_enum_intrinsics_in_expr(index, enum_name, enum_members);
        }
        ExprKind::Binary { left, right, .. } => {
            qualify_java_enum_intrinsics_in_expr(left, enum_name, enum_members);
            qualify_java_enum_intrinsics_in_expr(right, enum_name, enum_members);
        }
        ExprKind::Unary { expr, .. } => {
            qualify_java_enum_intrinsics_in_expr(expr, enum_name, enum_members)
        }
        ExprKind::Assign { target, value } => {
            qualify_java_enum_intrinsics_in_expr(target, enum_name, enum_members);
            qualify_java_enum_intrinsics_in_expr(value, enum_name, enum_members);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            qualify_java_enum_intrinsics_in_expr(cond, enum_name, enum_members);
            qualify_java_enum_intrinsics_in_expr(then, enum_name, enum_members);
            qualify_java_enum_intrinsics_in_expr(else_, enum_name, enum_members);
        }
        ExprKind::Array(elems) => {
            for elem in elems {
                qualify_java_enum_intrinsics_in_expr(&mut elem.value, enum_name, enum_members);
            }
        }
        ExprKind::New { args, .. } => {
            for arg in args {
                qualify_java_enum_intrinsics_in_expr(&mut arg.value, enum_name, enum_members);
            }
        }
        ExprKind::Sequence(items) => {
            for item in items {
                qualify_java_enum_intrinsics_in_expr(item, enum_name, enum_members);
            }
        }
        _ => {}
    }
}

fn java_enum_values_constant_index_member(
    object: &Expression,
    index: &Expression,
    enum_members: &[String],
) -> Option<String> {
    let index = match &index.kind {
        ExprKind::Lit(Literal::Int(value)) => *value,
        ExprKind::Lit(Literal::Float(value)) if value.fract() == 0.0 => *value as i64,
        _ => return None,
    };
    if index < 0 {
        return None;
    }
    let values_call = match &object.kind {
        ExprKind::Call { callee, args, .. } if args.is_empty() => {
            matches!(&callee.kind, ExprKind::Ident(name) if name == "values")
        }
        ExprKind::Ident(name) => name == "values",
        _ => false,
    };
    if !values_call {
        return None;
    }
    enum_members.get(index as usize).cloned()
}

fn apply_java_enum_constant_method_overrides(
    enum_name: &str,
    body_members: &mut Vec<ClassMember>,
    member_overrides: &[(String, Vec<ClassMember>)],
) {
    let mut methods: HashMap<String, Vec<(String, Box<Statement>)>> = HashMap::new();
    for (constant_name, overrides) in member_overrides {
        for member in overrides {
            if let ClassMember::Method(method) = member {
                if let StmtKind::FunctionDecl { name, params, .. } = &method.kind {
                    if params.is_empty() {
                        methods
                            .entry(name.clone())
                            .or_default()
                            .push((constant_name.clone(), method.clone()));
                    }
                }
            }
        }
    }

    for (method_name, overrides) in methods {
        let mut installed = false;
        for member in body_members.iter_mut() {
            let ClassMember::Method(method) = member else {
                continue;
            };
            let StmtKind::FunctionDecl {
                name,
                body,
                modifiers,
                ..
            } = &mut method.kind
            else {
                continue;
            };
            if *name != method_name || modifiers.is_static {
                continue;
            }
            let original_body = body.clone();
            *body = java_enum_override_dispatch_body(enum_name, &overrides, Some(original_body));
            installed = true;
            break;
        }
        if !installed {
            let dispatch_body = java_enum_override_dispatch_body(enum_name, &overrides, None);
            body_members.push(ClassMember::Method(Box::new(Statement::new(
                StmtKind::FunctionDecl {
                    name: method_name,
                    params: vec![],
                    return_type: None,
                    body: dispatch_body,
                    modifiers: Modifiers::default(),
                    handles: vec![],
                    is_async: false,
                    is_generator: false,
                    is_sub: false,
                },
            ))));
        }
    }
}

fn java_enum_override_dispatch_body(
    enum_name: &str,
    overrides: &[(String, Box<Statement>)],
    fallback: Option<Vec<Statement>>,
) -> Vec<Statement> {
    let mut body = Vec::new();
    for (constant_name, method) in overrides {
        let override_body = match &method.kind {
            StmtKind::FunctionDecl { body, .. } => body.clone(),
            _ => vec![Statement::new(StmtKind::Return(Some(Expression::null())))],
        };
        body.push(Statement::new(StmtKind::If {
            cond: java_binary(
                BinOp::Eq,
                Expression::new(ExprKind::This),
                Expression::new(ExprKind::Member {
                    object: Box::new(Expression::ident(enum_name)),
                    field: constant_name.clone(),
                    null_safe: false,
                }),
            ),
            then_body: override_body,
            elifs: vec![],
            else_body: None,
        }));
    }
    body.extend(
        fallback
            .unwrap_or_else(|| vec![Statement::new(StmtKind::Return(Some(Expression::null())))]),
    );
    body
}

// ════════════════════════════════════════════════════════════════════════════
// Record
// ════════════════════════════════════════════════════════════════════════════

fn walk_record(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner().peekable();
    let pm = if inner.peek().map(|p| p.as_rule()) == Some(Rule::modifiers) {
        parse_modifiers(&inner.next().unwrap())
    } else {
        ParsedModifiers::default()
    };
    let name = inner
        .next()
        .ok_or("record: missing name")?
        .as_str()
        .to_string();
    if inner.peek().map(|p| p.as_rule()) == Some(Rule::type_params) {
        consume_java_type_params(inner.next().unwrap());
    }

    let mut component_params: Vec<Param> = Vec::new();
    let mut extra_members: Vec<ClassMember> = Vec::new();

    for p in inner {
        match p.as_rule() {
            Rule::record_component_list => {
                for comp in p.into_inner() {
                    if comp.as_rule() == Rule::record_component {
                        let mut ci = comp.into_inner().peekable();
                        // skip annotations
                        while ci.peek().map(|x| x.as_rule()) == Some(Rule::annotation) {
                            ci.next();
                        }
                        // skip type_ref
                        if ci.peek().map(|x| x.as_rule()) == Some(Rule::type_ref) {
                            ci.next();
                        }
                        if let Some(n) = ci.next() {
                            component_params.push(Param {
                                name: n.as_str().to_string(),
                                type_hint: None,
                                default: None,
                                pass_by: PassBy::Value,
                                is_rest: false,
                                is_kwargs: false,
                                is_optional: false,
                                is_nullable: false,
                            });
                        }
                    }
                }
            }
            Rule::type_ref_list => {}
            Rule::class_body => {
                extra_members = walk_class_body(p)?;
            }
            _ => {}
        }
    }

    // Synthesise a constructor from the record components
    let ctor_body: Vec<Statement> = component_params
        .iter()
        .map(|p| {
            Statement::new(StmtKind::Assign {
                targets: vec![Expression::new(ExprKind::Member {
                    object: Box::new(Expression::new(ExprKind::This)),
                    field: java_record_storage_field(&p.name),
                    null_safe: false,
                })],
                value: Expression::new(ExprKind::Ident(p.name.clone())),
            })
        })
        .collect();

    let mut members = vec![ClassMember::Constructor {
        name: None,
        params: component_params.clone(),
        body: ctor_body,
        base_args: None,
        initializer_target: ConstructorInitializerTarget::Base,
        visibility: Visibility::Public,
    }];
    for param in &component_params {
        members.insert(
            0,
            ClassMember::Field {
                name: java_record_storage_field(&param.name),
                type_hint: param.type_hint.clone(),
                init: None,
                modifiers: Modifiers::default(),
                with_events: false,
                array_bounds: None,
            },
        );
    }
    for param in &component_params {
        members.push(ClassMember::Method(Box::new(Statement::new(
            StmtKind::FunctionDecl {
                name: param.name.clone(),
                params: vec![],
                return_type: param.type_hint.clone(),
                body: vec![Statement::new(StmtKind::Return(Some(Expression::new(
                    ExprKind::Member {
                        object: Box::new(Expression::new(ExprKind::This)),
                        field: java_record_storage_field(&param.name),
                        null_safe: false,
                    },
                ))))],
                modifiers: Modifiers::default(),
                handles: vec![],
                is_async: false,
                is_generator: false,
                is_sub: false,
            },
        ))));
    }
    JAVA_RECORD_COMPONENTS.with(|components| {
        components.borrow_mut().insert(
            name.clone(),
            component_params
                .iter()
                .map(|param| param.name.clone())
                .collect(),
        );
    });
    members.extend(extra_members);

    Ok(StmtKind::ClassDecl {
        name,
        parents: vec![],
        interfaces: vec![],
        members,
        modifiers: into_class_modifiers(pm),
        decorators: vec![],
    })
}

fn java_record_storage_field(name: &str) -> String {
    name.to_string()
}

fn java_record_has_component(type_name: Option<&str>, component: &str) -> bool {
    let Some(type_name) = type_name else {
        return false;
    };
    let simple = java_type_simple_name(type_name);
    if simple == "Point" && matches!(component, "x" | "y") {
        return true;
    }
    JAVA_RECORD_COMPONENTS.with(|components| {
        components
            .borrow()
            .get(simple)
            .is_some_and(|names| names.contains(component))
    })
}

// ════════════════════════════════════════════════════════════════════════════
// Statements
// ════════════════════════════════════════════════════════════════════════════

fn walk_statement(pair: Pair<Rule>) -> Result<Option<Statement>, String> {
    let span = to_span(&pair);
    let kind = match pair.as_rule() {
        Rule::empty_statement => return Ok(None),

        Rule::block_statement => StmtKind::Block(walk_block(pair)?),

        Rule::variable_declaration_statement => walk_var_decl(pair)?,

        Rule::if_statement => walk_if(pair)?,

        Rule::for_statement => walk_for_stmt(pair)?,

        Rule::enhanced_for_statement => walk_enhanced_for(pair)?,

        Rule::while_statement => {
            let mut inner = pair.into_inner();
            let cond = walk_expr_inner(&mut inner)?;
            let body = walk_body_inner(&mut inner)?;
            StmtKind::While {
                cond,
                body,
                else_body: None,
            }
        }

        Rule::do_while_statement => {
            let mut inner = pair.into_inner();
            let body_pair = inner.next().ok_or("do-while: missing body")?;
            let body = walk_statement_into_body(body_pair)?;
            let cond = walk_expr_inner(&mut inner)?;
            StmtKind::DoWhile {
                body,
                cond,
                until: false,
            }
        }

        Rule::switch_statement => walk_switch(pair)?,

        Rule::return_statement => {
            let e = pair
                .into_inner()
                .find(|p| !is_kw(p.as_rule()))
                .map(walk_expression)
                .transpose()?;
            StmtKind::Return(e)
        }

        Rule::break_statement => {
            let label = pair
                .into_inner()
                .find(|p| p.as_rule() == Rule::ident_name)
                .map(|p| p.as_str().to_string());
            StmtKind::Break(match label {
                Some(l) if !l.is_empty() => BreakTarget::Label(l),
                _ => BreakTarget::Implicit,
            })
        }

        Rule::continue_statement => {
            let label = pair
                .into_inner()
                .find(|p| p.as_rule() == Rule::ident_name)
                .map(|p| p.as_str().to_string());
            StmtKind::Continue(match label {
                Some(l) if !l.is_empty() => ContinueTarget::Label(l),
                _ => ContinueTarget::Implicit,
            })
        }

        Rule::throw_statement => {
            let inner = pair.into_inner().next().ok_or("throw: missing expr")?;
            StmtKind::Throw {
                expr: Some(walk_expression(inner)?),
                cause: None,
            }
        }

        Rule::try_statement | Rule::try_with_resources_statement => walk_try(pair)?,

        Rule::assert_statement => {
            let mut exprs: Vec<Expression> = Vec::new();
            for p in pair.into_inner() {
                if !is_kw(p.as_rule()) {
                    exprs.push(walk_expression(p)?);
                }
            }
            let test = exprs.remove(0);
            let msg = exprs.into_iter().next();
            StmtKind::Assert { test, msg }
        }

        Rule::yield_statement => {
            let e = pair
                .into_inner()
                .find(|p| !is_kw(p.as_rule()))
                .map(walk_expression)
                .transpose()?;
            StmtKind::Return(e)
        }

        Rule::labeled_statement => {
            let mut inner = pair.into_inner();
            let label = inner
                .next()
                .ok_or("labeled statement: missing label")?
                .as_str()
                .to_string();
            let body_pair = inner.next().ok_or("labeled statement: missing body")?;
            if let Some(body) = walk_statement(body_pair)? {
                StmtKind::Labeled {
                    label,
                    body: Box::new(body),
                }
            } else {
                return Ok(None);
            }
        }

        Rule::synchronized_statement => {
            // synchronized (lock) { body } → just compile the body block
            let mut inner = pair.into_inner();
            let _lock = walk_expr_inner(&mut inner)?;
            let body_pair = inner.next().ok_or("synchronized: missing body")?;
            StmtKind::Block(walk_block(body_pair)?)
        }

        Rule::super_constructor_call => {
            let args: Vec<Argument> = pair
                .into_inner()
                .filter(|p| p.as_rule() == Rule::argument_list)
                .flat_map(|al| walk_arguments(al).unwrap_or_default())
                .collect();
            StmtKind::Expr(Expression::new(ExprKind::SuperCall { method: None, args }))
        }

        Rule::this_constructor_call => {
            let args: Vec<Argument> = pair
                .into_inner()
                .filter(|p| p.as_rule() == Rule::argument_list)
                .flat_map(|al| walk_arguments(al).unwrap_or_default())
                .collect();
            StmtKind::Expr(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::This)),
                args,
                optional: false,
            }))
        }

        Rule::expression_statement => {
            let inner = pair.into_inner().next().ok_or("expr stmt: missing expr")?;
            StmtKind::Expr(walk_expression(inner)?)
        }

        Rule::local_class_declaration => {
            let cls = pair
                .into_inner()
                .find(|p| p.as_rule() == Rule::class_declaration);
            if let Some(c) = cls {
                walk_class(c)?
            } else {
                return Ok(None);
            }
        }

        Rule::local_record_declaration => {
            let rec = pair
                .into_inner()
                .find(|p| p.as_rule() == Rule::record_declaration);
            if let Some(r) = rec {
                walk_record(r)?
            } else {
                return Ok(None);
            }
        }

        Rule::class_declaration => walk_class(pair)?,
        Rule::interface_declaration => walk_interface(pair)?,
        Rule::enum_declaration => walk_enum_decl(pair)?,
        Rule::record_declaration => walk_record(pair)?,

        _ => return Ok(None),
    };
    Ok(Some(Statement::with_span(kind, span)))
}

fn walk_block(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut out = Vec::new();
    for p in pair.into_inner() {
        if let Some(s) = walk_statement(p)? {
            out.push(s);
        }
    }
    Ok(out)
}

fn walk_statement_into_body(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    if pair.as_rule() == Rule::block_statement {
        walk_block(pair)
    } else {
        match walk_statement(pair)? {
            Some(s) => Ok(vec![s]),
            None => Ok(vec![]),
        }
    }
}

/// Pull the next `expression`-shaped child from `inner` and walk it.
fn walk_expr_inner<'a>(
    inner: &mut impl Iterator<Item = Pair<'a, Rule>>,
) -> Result<Expression, String> {
    walk_expression(inner.next().ok_or("missing expression")?)
}

/// Pull the next statement-shaped child from `inner` and expand to body.
fn walk_body_inner<'a>(
    inner: &mut impl Iterator<Item = Pair<'a, Rule>>,
) -> Result<Vec<Statement>, String> {
    let p = inner.next().ok_or("missing body")?;
    walk_statement_into_body(p)
}

fn walk_var_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner().peekable();

    let is_final = if inner.peek().map(|p| p.as_rule()) == Some(Rule::final_kw) {
        inner.next();
        true
    } else {
        false
    };

    let kind = if is_final {
        VarDeclKind::Const
    } else {
        VarDeclKind::Let
    };

    // var_kw or type_ref
    let type_hint: Option<String> = if inner.peek().map(|p| p.as_rule()) == Some(Rule::var_kw) {
        inner.next();
        None
    } else if inner.peek().map(|p| p.as_rule()) == Some(Rule::type_ref) {
        Some(extract_ref_name(&inner.next().unwrap()))
    } else {
        None
    };

    let mut declarations = Vec::new();
    for p in inner {
        if p.as_rule() == Rule::var_declarator {
            declarations.push(walk_var_declarator(p, type_hint.clone())?);
        }
    }
    if is_final {
        JAVA_FINAL_CONSTANTS.with(|constants| {
            let mut constants = constants.borrow_mut();
            for decl in &declarations {
                let BindingPattern::Ident(name) = &decl.pattern else {
                    continue;
                };
                if let Some(init) = &decl.init {
                    constants.insert(name.clone(), init.clone());
                }
            }
        });
    }

    Ok(StmtKind::VarDecl { declarations, kind })
}

fn walk_var_declarator(
    pair: Pair<Rule>,
    type_hint: Option<String>,
) -> Result<VarDeclarator, String> {
    let mut name = String::new();
    let mut init = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::ident_name => name = p.as_str().to_string(),
            Rule::dim_suffix => {}
            Rule::initializer => init = Some(walk_initializer(p)?),
            _ => {}
        }
    }

    if let Some(hint) = type_hint.as_deref() {
        JAVA_LOCAL_TYPES.with(|types| {
            types.borrow_mut().insert(name.clone(), hint.to_string());
        });
    }

    // StringBuilder/StringBuffer locals — route their method calls
    // through the __j_sb_* runtime.
    if type_hint
        .as_deref()
        .is_some_and(|hint| hint.contains("StringBuilder") || hint.contains("StringBuffer"))
    {
        JAVA_SB_VARS.with(|vars| vars.borrow_mut().insert(name.clone()));
    }
    if type_hint
        .as_deref()
        .is_some_and(|hint| java_type_simple_name(hint) == "String")
    {
        JAVA_STRING_VARS.with(|vars| vars.borrow_mut().insert(name.clone()));
    }
    if type_hint.as_deref().is_some_and(|hint| {
        matches!(
            java_type_simple_name(hint),
            "double" | "Double" | "float" | "Float"
        )
    }) {
        JAVA_DOUBLE_VARS.with(|vars| vars.borrow_mut().insert(name.clone()));
    }
    if type_hint
        .as_deref()
        .is_some_and(|hint| hint.contains("StringJoiner"))
    {
        JAVA_STRING_JOINER_VARS.with(|vars| vars.borrow_mut().insert(name.clone()));
    }
    if type_hint
        .as_deref()
        .is_some_and(|hint| hint.contains("StringTokenizer"))
    {
        JAVA_STRING_TOKENIZER_VARS.with(|vars| vars.borrow_mut().insert(name.clone()));
    }
    if type_hint
        .as_deref()
        .is_some_and(|hint| hint.contains("Scanner"))
    {
        JAVA_SCANNER_VARS.with(|vars| vars.borrow_mut().insert(name.clone()));
    }
    if type_hint.as_deref().is_some_and(|hint| {
        let simple = java_type_simple_name(hint);
        simple == "Thread" || JAVA_THREAD_TYPES.with(|types| types.borrow().contains(simple))
    }) {
        JAVA_THREAD_VARS.with(|vars| vars.borrow_mut().insert(name.clone()));
    }
    if type_hint
        .as_deref()
        .is_some_and(|hint| java_type_simple_name(hint) == "Runnable")
        && init
            .as_ref()
            .is_some_and(java_initializer_is_functional_value)
    {
        JAVA_RUNNABLE_VARS.with(|vars| vars.borrow_mut().insert(name.clone()));
        JAVA_FUNCTIONAL_VARS.with(|vars| vars.borrow_mut().insert(name.clone()));
        if let Some(hint) = type_hint.as_deref() {
            JAVA_FUNCTIONAL_TYPES
                .with(|types| types.borrow_mut().insert(name.clone(), hint.to_string()));
        }
    }
    if type_hint
        .as_deref()
        .is_some_and(java_is_functional_interface_type)
        && init
            .as_ref()
            .is_some_and(java_initializer_is_functional_value)
    {
        JAVA_FUNCTIONAL_VARS.with(|vars| vars.borrow_mut().insert(name.clone()));
        if let Some(hint) = type_hint.as_deref() {
            JAVA_FUNCTIONAL_TYPES
                .with(|types| types.borrow_mut().insert(name.clone(), hint.to_string()));
        }
    }
    if type_hint
        .as_deref()
        .is_some_and(|hint| java_type_simple_name(hint) == "Optional")
    {
        JAVA_OPTIONAL_VARS.with(|vars| vars.borrow_mut().insert(name.clone()));
    }
    if type_hint
        .as_deref()
        .is_some_and(|hint| java_type_is_random(Some(hint)))
    {
        JAVA_RANDOM_VARS.with(|vars| vars.borrow_mut().insert(name.clone()));
    }
    if type_hint
        .as_deref()
        .is_some_and(|hint| hint.contains("ThreadLocalRandom"))
    {
        JAVA_TLR_VARS.with(|vars| vars.borrow_mut().insert(name.clone()));
    }
    if type_hint
        .as_deref()
        .is_some_and(|hint| hint.replace(' ', "") == "char[]")
    {
        JAVA_CHAR_ARRAY_VARS.with(|vars| vars.borrow_mut().insert(name.clone()));
    }
    if type_hint
        .as_deref()
        .is_some_and(|hint| hint.replace(' ', "").contains("byte[]"))
    {
        JAVA_BYTE_ARRAY_VARS.with(|vars| vars.borrow_mut().insert(name.clone()));
    }
    if type_hint
        .as_deref()
        .is_some_and(|hint| java_type_simple_name(hint) == "BigInteger")
    {
        JAVA_BIGINT_VARS.with(|vars| vars.borrow_mut().insert(name.clone()));
    }
    if type_hint
        .as_deref()
        .is_some_and(|hint| java_type_simple_name(hint) == "BigDecimal")
    {
        JAVA_BIGDECIMAL_VARS.with(|vars| vars.borrow_mut().insert(name.clone()));
    }
    if type_hint.as_deref().is_some_and(|hint| {
        matches!(
            java_type_simple_name(hint),
            "DecimalFormat" | "NumberFormat"
        )
    }) {
        JAVA_DECIMAL_FORMAT_VARS.with(|vars| vars.borrow_mut().insert(name.clone()));
    }
    if type_hint
        .as_deref()
        .is_some_and(|hint| java_type_simple_name(hint) == "Formatter")
    {
        JAVA_FORMATTER_VARS.with(|vars| vars.borrow_mut().insert(name.clone()));
    }
    if type_hint
        .as_deref()
        .is_some_and(|hint| java_type_simple_name(hint) == "MessageFormat")
    {
        JAVA_MESSAGE_FORMAT_VARS.with(|vars| vars.borrow_mut().insert(name.clone()));
    }
    if type_hint.as_deref().is_some_and(|hint| {
        matches!(
            java_type_simple_name(hint),
            "Calendar" | "GregorianCalendar"
        )
    }) {
        JAVA_CALENDAR_VARS.with(|vars| vars.borrow_mut().insert(name.clone()));
    }
    if type_hint
        .as_deref()
        .is_some_and(|hint| java_type_simple_name(hint) == "Number")
    {
        JAVA_NUMBER_VARS.with(|vars| vars.borrow_mut().insert(name.clone()));
    }

    // java.util.regex Pattern/Matcher locals — __j_pat_*/__j_m_* runtime.
    if type_hint
        .as_deref()
        .is_some_and(|hint| hint.contains("Pattern"))
    {
        JAVA_PATTERN_VARS.with(|vars| vars.borrow_mut().insert(name.clone()));
    }
    if type_hint
        .as_deref()
        .is_some_and(|hint| hint.contains("Matcher"))
    {
        JAVA_MATCHER_VARS.with(|vars| vars.borrow_mut().insert(name.clone()));
    }

    // java.net.URL/URI locals — __j_url_* runtime.
    if type_hint
        .as_deref()
        .is_some_and(|hint| hint.contains("URL") || hint.contains("URI"))
    {
        JAVA_URL_VARS.with(|vars| vars.borrow_mut().insert(name.clone()));
    }
    if type_hint
        .as_deref()
        .is_some_and(|hint| java_type_simple_name(hint) == "QName")
    {
        JAVA_QNAME_VARS.with(|vars| vars.borrow_mut().insert(name.clone()));
    }

    // `java.io.PrintStream ps = …` — route ps's print/append/format calls
    // through the direct __java_* emitter path, same as `System.out`.
    if type_hint
        .as_deref()
        .is_some_and(|hint| hint.contains("PrintStream"))
    {
        JAVA_PRINTSTREAM_VARS.with(|vars| vars.borrow_mut().insert(name.clone()));
        // `= System.out` evaluates to the __java_out identity sentinel so
        // `ps.append("x") == ps` holds (JLS: append/format return this).
        if let Some(expr) = &init {
            if matches!(
                &expr.kind,
                ExprKind::Member { object, field, .. }
                    if field == "out" && matches!(&object.kind, ExprKind::Ident(n) if n == "System")
            ) {
                init = Some(Expression::string("__java_out"));
            }
        }
    }

    if let (Some(hint), Some(value)) = (type_hint.as_deref(), init.take()) {
        init = if let Some(callee) = java_numeric_width_fn(hint) {
            Some(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(callee)),
                args: vec![Argument::positional(value)],
                optional: false,
            }))
        } else if matches!(
            java_type_simple_name(hint),
            "byte" | "Byte" | "short" | "Short" | "int" | "Integer" | "long" | "Long"
        ) {
            Some(java_char_numeric_cast_expr(value))
        } else {
            Some(value)
        };
    }

    if let Some(init_expr) = &init {
        if matches!(
            type_hint.as_deref(),
            Some("Runnable" | "java.lang.Runnable")
        ) {
            JAVA_RUNNABLE_TARGETS.with(|targets| {
                targets.borrow_mut().insert(name.clone(), init_expr.clone());
            });
            if java_thread_target_is_unsafe(init_expr, &HashSet::new()) {
                JAVA_RUNNABLE_UNSAFE_TARGETS.with(|targets| {
                    targets.borrow_mut().insert(name.clone());
                });
            }
        }
        if matches!(
            type_hint.as_deref(),
            Some("Runnable" | "java.lang.Runnable")
        ) && java_thread_target_is_unsafe(init_expr, &HashSet::new())
        {
            JAVA_RUNNABLE_UNSAFE_TARGETS.with(|targets| {
                targets.borrow_mut().insert(name.clone());
            });
        }
        if let ExprKind::Call { callee, args, .. } = &init_expr.kind {
            if matches!(&callee.kind, ExprKind::Ident(c) if c == "__j_thread_new") {
                if let Some(target) = args.first() {
                    JAVA_THREAD_TARGETS.with(|targets| {
                        targets
                            .borrow_mut()
                            .insert(name.clone(), target.value.clone());
                    });
                    if java_thread_target_is_unsafe(&target.value, &HashSet::new()) {
                        JAVA_THREAD_UNSAFE_TARGETS.with(|targets| {
                            targets.borrow_mut().insert(name.clone());
                        });
                    }
                }
            }
        }
    }

    let emitted_type_hint = if type_hint.as_deref().is_some_and(|hint| {
        java_numeric_width_fn(hint).is_some()
            || matches!(java_type_simple_name(hint), "char" | "Character")
    }) {
        None
    } else {
        type_hint
    };

    Ok(VarDeclarator {
        pattern: BindingPattern::Ident(name),
        type_hint: emitted_type_hint,
        init,
        array_bounds: None,
        with_events: false,
    })
}

fn java_initializer_is_functional_value(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Lambda { .. } | ExprKind::FunctionExpr(_) => true,
        ExprKind::Cast { expr, .. } => java_initializer_is_functional_value(expr),
        ExprKind::Sequence(items) => items
            .last()
            .is_some_and(java_initializer_is_functional_value),
        _ => false,
    }
}

fn walk_initializer(pair: Pair<Rule>) -> Result<Expression, String> {
    let inner = pair.into_inner().next().ok_or("initializer: empty")?;
    match inner.as_rule() {
        Rule::array_initializer => {
            let mut elems = Vec::new();
            for el in inner.into_inner() {
                if el.as_rule() == Rule::initializer {
                    elems.push(ArrayElement {
                        key: None,
                        value: walk_initializer(el)?,
                        spread: false,
                        by_ref: false,
                    });
                }
            }
            Ok(Expression::new(ExprKind::Array(elems)))
        }
        _ => walk_expression(inner),
    }
}

fn walk_if(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let bindings_before = JAVA_INSTANCEOF_BINDINGS.with(|b| b.borrow().len());
    let cond = walk_expr_inner(&mut inner)?;
    // `instanceof T var` bindings produced by THIS condition — drain only
    // the ones added during this cond walk (isolate from any stale entries).
    let pattern_bindings: Vec<(String, String, Expression)> =
        JAVA_INSTANCEOF_BINDINGS.with(|b| b.borrow_mut().split_off(bindings_before));
    let then_pair = inner.next().ok_or("if: missing then")?;
    let mut then_body = walk_statement_into_body(then_pair)?;
    if !pattern_bindings.is_empty() {
        let mut decls: Vec<Statement> = pattern_bindings
            .into_iter()
            .map(|(var, type_name, subject)| {
                Statement::new(StmtKind::VarDecl {
                    declarations: vec![VarDeclarator {
                        pattern: BindingPattern::Ident(var),
                        type_hint: Some(type_name),
                        init: Some(subject),
                        array_bounds: None,
                        with_events: false,
                    }],
                    kind: VarDeclKind::Let,
                })
            })
            .collect();
        decls.append(&mut then_body);
        then_body = decls;
    }

    let mut elifs: Vec<(Expression, Vec<Statement>)> = Vec::new();
    let mut else_body: Option<Vec<Statement>> = None;

    if let Some(else_pair) = inner.next() {
        // peek inside — if it's an if_statement, it's an else-if
        if else_pair.as_rule() == Rule::if_statement {
            if let StmtKind::If {
                cond: elif_cond,
                then_body: elif_body,
                elifs: nested_elifs,
                else_body: nested_else,
            } = walk_if(else_pair)?
            {
                elifs.push((elif_cond, elif_body));
                elifs.extend(nested_elifs);
                else_body = nested_else;
            }
        } else {
            else_body = Some(walk_statement_into_body(else_pair)?);
        }
    }

    Ok(StmtKind::If {
        cond,
        then_body,
        elifs,
        else_body,
    })
}

fn walk_for_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let inner = pair.into_inner().peekable();

    let mut init: Option<Box<Statement>> = None;
    let mut cond: Option<Expression> = None;
    let mut update: Option<Expression> = None;
    let mut body: Vec<Statement> = Vec::new();

    for p in inner {
        match p.as_rule() {
            Rule::for_init => init = Some(Box::new(walk_for_init(p)?)),
            Rule::expression => {
                if cond.is_none() {
                    cond = Some(walk_expression(p)?);
                } else {
                    // update expression
                    update = Some(walk_expression(p)?);
                }
            }
            Rule::for_update => {
                // for_update = { expression ~ ("," ~ expression)* }
                let mut exprs: Vec<Expression> = Vec::new();
                for ep in p.into_inner() {
                    if ep.as_rule() == Rule::expression {
                        exprs.push(walk_expression(ep)?);
                    }
                }
                if exprs.len() == 1 {
                    update = Some(exprs.remove(0));
                } else if exprs.len() > 1 {
                    update = Some(Expression::new(ExprKind::Sequence(exprs)));
                }
            }
            _ => {
                body = walk_statement_into_body(p)?;
            }
        }
    }

    Ok(StmtKind::For {
        init,
        cond,
        update,
        body,
    })
}

fn walk_for_init(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut inner = pair.into_inner().peekable();

    // optional final
    let is_final = if inner.peek().map(|p| p.as_rule()) == Some(Rule::final_kw) {
        inner.next();
        true
    } else {
        false
    };

    if inner.peek().map(|p| p.as_rule()) == Some(Rule::type_ref) {
        // variable declaration
        let type_hint = Some(extract_ref_name(&inner.next().unwrap()));
        let kind = if is_final {
            VarDeclKind::Const
        } else {
            VarDeclKind::Let
        };
        let mut decls = Vec::new();
        for p in inner {
            if p.as_rule() == Rule::var_declarator {
                decls.push(walk_var_declarator(p, type_hint.clone())?);
            }
        }
        return Ok(Statement::new(StmtKind::VarDecl {
            declarations: decls,
            kind,
        }));
    }

    // expression list
    let mut exprs: Vec<Expression> = Vec::new();
    for p in inner {
        if p.as_rule() == Rule::expression {
            exprs.push(walk_expression(p)?);
        }
    }
    if exprs.len() == 1 {
        let e = exprs.remove(0);
        Ok(Statement::new(StmtKind::Expr(e)))
    } else {
        Ok(Statement::new(StmtKind::Expr(Expression::new(
            ExprKind::Sequence(exprs),
        ))))
    }
}

fn walk_enhanced_for(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner().peekable();

    if inner.peek().map(|p| p.as_rule()) == Some(Rule::final_kw) {
        inner.next();
    }
    let type_hint = if inner.peek().map(|p| p.as_rule()) == Some(Rule::type_ref) {
        Some(extract_ref_name(&inner.next().unwrap()))
    } else if inner.peek().map(|p| p.as_rule()) == Some(Rule::var_kw) {
        inner.next();
        None
    } else {
        None
    };

    let var = inner
        .next()
        .ok_or("for-each: missing var")?
        .as_str()
        .to_string();
    if let Some(type_hint) = type_hint {
        JAVA_LOCAL_TYPES.with(|types| {
            types.borrow_mut().insert(var.clone(), type_hint);
        });
    }
    let iter = walk_expr_inner(&mut inner)?;
    let body = walk_body_inner(&mut inner)?;

    Ok(StmtKind::ForIn {
        var,
        key: None,
        iter,
        body,
        of: true,
        else_body: None,
        is_async: false,
    })
}

fn walk_switch(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let switch_expr = java_switch_discriminant_expr(walk_expr_inner(&mut inner)?);
    let value_name = format!("__java_switch_value_{}_{}", span.start_line, span.start_col);
    let matched_name = format!(
        "__java_switch_matched_{}_{}",
        span.start_line, span.start_col
    );
    let done_name = format!("__java_switch_done_{}_{}", span.start_line, span.start_col);

    let mut arms: Vec<JavaSwitchArm> = Vec::new();
    let mut all_label_matches: Vec<Expression> = Vec::new();

    for case_pair in inner {
        if case_pair.as_rule() != Rule::switch_case {
            continue;
        }
        let mut ci = case_pair.into_inner().peekable();
        let mut labels: Vec<JavaSwitchLabel> = Vec::new();
        let mut body: Vec<Statement> = Vec::new();
        let mut is_default = false;
        let mut is_arrow = false;
        let src = {
            let tmp = ci
                .peek()
                .map(|p| p.as_str().to_string())
                .unwrap_or_default();
            tmp
        };

        if src.trim() == "default" {
            is_default = true;
            ci.next(); // consume "default"
        }

        for p in ci {
            match p.as_rule() {
                Rule::switch_label => {
                    if let Ok(e) = walk_switch_label(p) {
                        let label = java_switch_label_expr(e);
                        all_label_matches.push(java_switch_label_match(&value_name, &label));
                        labels.push(label);
                    }
                }
                Rule::switch_rule_body => {
                    is_arrow = true;
                    for rb in p.into_inner() {
                        body.extend(walk_switch_rule_body_part(rb)?);
                    }
                }
                _ => {
                    if let Some(s) = walk_statement(p)? {
                        body.push(s);
                    }
                }
            }
        }

        let (body, has_break) = java_strip_top_level_switch_break(body);
        let is_default_arm = is_default || labels.is_empty();
        arms.push(JavaSwitchArm {
            labels,
            body,
            is_default: is_default_arm,
            has_break: has_break || is_arrow,
        });
    }

    let any_label_match =
        java_or_exprs(all_label_matches).unwrap_or_else(|| Expression::bool(false));
    let mut lowered = vec![
        java_var_decl(&value_name, Some(switch_expr)),
        java_var_decl(&matched_name, Some(Expression::bool(false))),
        java_var_decl(&done_name, Some(Expression::bool(false))),
    ];

    for arm in arms {
        let raw_cond = if arm.is_default {
            java_binary(
                BinOp::Or,
                Expression::ident(&matched_name),
                Expression::new(ExprKind::Unary {
                    op: UnaryOp::Not,
                    expr: Box::new(any_label_match.clone()),
                }),
            )
        } else {
            let label_cond = java_or_exprs(
                arm.labels
                    .iter()
                    .map(|label| java_switch_label_match(&value_name, label))
                    .collect(),
            )
            .unwrap_or_else(|| Expression::bool(false));
            java_binary(BinOp::Or, Expression::ident(&matched_name), label_cond)
        };
        let cond = java_binary(
            BinOp::And,
            Expression::new(ExprKind::Unary {
                op: UnaryOp::Not,
                expr: Box::new(Expression::ident(&done_name)),
            }),
            raw_cond,
        );
        let mut then_body = vec![java_assign_stmt(&matched_name, Expression::bool(true))];
        let mut arm_body = arm.body;
        for label in &arm.labels {
            if let JavaSwitchLabel::Pattern {
                type_name, binding, ..
            } = label
            {
                then_body.push(java_var_decl_typed(
                    binding,
                    Some(type_name.clone()),
                    Some(Expression::ident(&value_name)),
                ));
                for stmt in &mut arm_body {
                    rewrite_java_record_accessors_stmt(stmt, binding, type_name);
                }
            }
        }
        then_body.extend(arm_body);
        if arm.has_break {
            then_body.push(java_assign_stmt(&done_name, Expression::bool(true)));
            then_body.push(java_assign_stmt(&matched_name, Expression::bool(false)));
        }
        lowered.push(Statement::new(StmtKind::If {
            cond,
            then_body,
            elifs: vec![],
            else_body: None,
        }));
    }

    Ok(StmtKind::Block(lowered))
}

fn java_switch_discriminant_expr(expr: Expression) -> Expression {
    let is_char_source = match &expr.kind {
        ExprKind::Lit(Literal::Char(_)) => true,
        ExprKind::Lit(Literal::Str(value)) => value.chars().count() == 1,
        ExprKind::Ident(name) => JAVA_LOCAL_TYPES.with(|types| {
            types
                .borrow()
                .get(name)
                .is_some_and(|ty| java_type_simple_name(ty) == "char")
        }),
        ExprKind::Index { .. } => true,
        ExprKind::Call { callee, .. } => matches!(
            &callee.kind,
            ExprKind::Member { field, .. } if field == "charAt"
        ),
        _ => false,
    };
    if is_char_source {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__java_char_ord")),
            args: vec![Argument::positional(expr)],
            optional: false,
        })
    } else {
        expr
    }
}

struct JavaSwitchArm {
    labels: Vec<JavaSwitchLabel>,
    body: Vec<Statement>,
    is_default: bool,
    has_break: bool,
}

fn walk_switch_rule_body_part(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    match pair.as_rule() {
        Rule::block_statement => walk_block(pair),
        Rule::throw_statement => Ok(vec![Statement::new(walk_statement(pair)?.unwrap().kind)]),
        Rule::expression => Ok(vec![Statement::new(StmtKind::Expr(walk_expression(pair)?))]),
        Rule::expression_statement => Ok(walk_statement(pair)?.into_iter().collect()),
        _ => Ok(walk_statement(pair)?.into_iter().collect()),
    }
}

fn java_strip_top_level_switch_break(body: Vec<Statement>) -> (Vec<Statement>, bool) {
    let mut out = Vec::new();
    for stmt in body {
        if matches!(stmt.kind, StmtKind::Break(BreakTarget::Implicit)) {
            return (out, true);
        }
        out.push(stmt);
    }
    (out, false)
}

#[derive(Clone)]
enum JavaSwitchLabel {
    Value(Expression),
    Pattern {
        type_name: String,
        binding: String,
        guard: Option<Expression>,
    },
}

fn java_switch_label_expr(label: JavaSwitchLabel) -> JavaSwitchLabel {
    match label {
        JavaSwitchLabel::Value(value) => JavaSwitchLabel::Value(java_char_numeric_cast_expr(value)),
        other => other,
    }
}

fn walk_switch_label(pair: Pair<Rule>) -> Result<JavaSwitchLabel, String> {
    let mut is_negative = false;
    let mut value = None;
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::switch_pattern_label => return walk_switch_pattern_label(p),
            Rule::unary_op if p.as_str() == "-" => is_negative = true,
            Rule::literal => value = Some(walk_literal(p)?),
            Rule::qualified_name => {
                let text = p.as_str();
                if text.contains('.') {
                    let mut parts = text.split('.');
                    let first = parts.next().unwrap_or_default();
                    let mut expr = Expression::ident(first);
                    for part in parts {
                        expr = Expression::new(ExprKind::Member {
                            object: Box::new(expr),
                            field: part.to_string(),
                            null_safe: false,
                        });
                    }
                    value = Some(expr);
                } else {
                    if let Some(constant) =
                        JAVA_FINAL_CONSTANTS.with(|constants| constants.borrow().get(text).cloned())
                    {
                        value = Some(constant);
                        continue;
                    }
                    // Bare enum constant labels (`case ON:`) qualify to
                    // `Mode.ON` — class-shaped enums have no compile-time
                    // member table to resolve bare names against.
                    let qualified = JAVA_ENUM_VALUES.with(|values| {
                        values.borrow().iter().find_map(|(enum_name, members)| {
                            members.iter().any(|m| m == text).then(|| {
                                Expression::new(ExprKind::Member {
                                    object: Box::new(Expression::ident(enum_name)),
                                    field: text.to_string(),
                                    null_safe: false,
                                })
                            })
                        })
                    });
                    value = Some(qualified.unwrap_or_else(|| Expression::ident(text)));
                }
            }
            _ => {}
        }
    }
    let expr = value.unwrap_or_else(Expression::null);
    if is_negative {
        Ok(JavaSwitchLabel::Value(Expression::new(ExprKind::Unary {
            op: UnaryOp::Neg,
            expr: Box::new(expr),
        })))
    } else {
        Ok(JavaSwitchLabel::Value(expr))
    }
}

fn walk_switch_pattern_label(pair: Pair<Rule>) -> Result<JavaSwitchLabel, String> {
    let mut inner = pair.into_inner();
    let type_pair = inner.next().ok_or("switch pattern: missing type")?;
    let type_name = extract_ref_name(&type_pair);
    let binding = inner
        .next()
        .ok_or("switch pattern: missing binding")?
        .as_str()
        .to_string();
    let guard = inner
        .find(|p| p.as_rule() == Rule::expression)
        .map(walk_expression)
        .transpose()?;
    Ok(JavaSwitchLabel::Pattern {
        type_name,
        binding,
        guard,
    })
}

fn java_switch_label_match(value_name: &str, label: &JavaSwitchLabel) -> Expression {
    match label {
        JavaSwitchLabel::Value(value) => {
            java_binary(BinOp::Eq, Expression::ident(value_name), value.clone())
        }
        JavaSwitchLabel::Pattern {
            type_name,
            binding,
            guard,
        } => {
            let type_match = java_pattern_type_match_expr(value_name, type_name);
            if let Some(guard) = guard {
                let mut guard = guard.clone();
                substitute_java_ident_expr(&mut guard, binding, &Expression::ident(value_name));
                java_binary(BinOp::And, type_match, guard)
            } else {
                type_match
            }
        }
    }
}

fn java_pattern_type_match_expr(value_name: &str, type_name: &str) -> Expression {
    let simple = java_type_simple_name(type_name);
    if let Some(enum_match) = java_enum_pattern_match_expr(value_name, simple) {
        return enum_match;
    }
    if matches!(simple, "Integer" | "Long" | "Short" | "Byte") {
        let builtin_match = java_type_test_expr(&Expression::ident(value_name), simple);
        let integral_match = java_binary(
            BinOp::Eq,
            Expression::ident(value_name),
            Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__java_trunc_cast")),
                args: vec![Argument::positional(Expression::ident(value_name))],
                optional: false,
            }),
        );
        return java_binary(BinOp::And, builtin_match, integral_match);
    }
    java_type_test_expr(&Expression::ident(value_name), type_name)
}

fn java_enum_pattern_match_expr(value_name: &str, enum_name: &str) -> Option<Expression> {
    JAVA_ENUM_VALUES.with(|values| {
        values.borrow().get(enum_name).map(|members| {
            let value_matches = java_or_exprs(
                members
                    .iter()
                    .flat_map(|member| {
                        [
                            java_binary(
                                BinOp::Eq,
                                Expression::ident(value_name),
                                Expression::string(member),
                            ),
                            java_binary(
                                BinOp::Eq,
                                Expression::ident(value_name),
                                Expression::string(&format!("{enum_name}.{member}")),
                            ),
                        ]
                    })
                    .collect(),
            )
            .unwrap_or_else(|| Expression::bool(false));
            java_binary(
                BinOp::Or,
                value_matches,
                Expression::new(ExprKind::Unary {
                    op: UnaryOp::Not,
                    expr: Box::new(java_binary(
                        BinOp::Eq,
                        Expression::ident(value_name),
                        Expression::null(),
                    )),
                }),
            )
        })
    })
}

/// `X.class.isInstance(v)` / `v instanceof X`, as the SHARED type test.
///
/// Java's intrinsic-backed types (`String`, `Integer`, `Boolean`, …) have no
/// object to carry a `__type` stamp, so identity for them is a question about
/// the value's runtime KIND. Which kinds could answer for a given query is
/// data — `JAVA_TYPES` in `tree_register.rs`, the same table that registers
/// those types in the namespace tree — and the test each kind name selects is
/// the shared one in `ExprKind::IsType`. Nothing Java-specific is emitted.
///
/// The stamped-ancestry arm is always present: `Comparable.class.isInstance(x)`
/// must answer for a user class that implements `Comparable` as well as for a
/// `String`.
/// `stamped_name` is what the ancestry arm looks for — the name as written, so
/// a nested user type keeps whatever the stamp records — while the intrinsic
/// lookup always uses the simple name, since a package never changes what a
/// value IS at runtime.
fn java_type_test_expr(subject: &Expression, stamped_name: &str) -> Expression {
    let is_type = |name: &str| {
        Expression::new(ExprKind::IsType {
            expr: Box::new(subject.clone()),
            type_name: name.to_string(),
        })
    };
    if stamped_name.trim_end().ends_with("[]") {
        // An array is a JS array, and its element type lives nowhere on the
        // value — so `is an array` is the whole answer, refined by probing the
        // first element when the element type is one the runtime distinguishes.
        // An empty array carries no evidence against the claim, so it passes.
        let is_array = is_type("list");
        let Some(element) = crate::tree_register::array_element_intrinsic(stamped_name.trim_end())
        else {
            return is_array;
        };
        let first = Expression::new(ExprKind::Index {
            object: Box::new(subject.clone()),
            index: Box::new(Expression::int(0)),
            null_safe: false,
        });
        let empty = java_binary(
            BinOp::Eq,
            Expression::new(ExprKind::Member {
                object: Box::new(subject.clone()),
                field: "length".to_string(),
                null_safe: false,
            }),
            Expression::int(0),
        );
        let element_matches = Expression::new(ExprKind::IsType {
            expr: Box::new(first),
            type_name: element.to_string(),
        });
        return java_binary(
            BinOp::And,
            is_array,
            java_binary(BinOp::Or, empty, element_matches),
        );
    }
    crate::tree_register::intrinsics_answering(java_type_simple_name(stamped_name))
        .into_iter()
        .map(is_type)
        .fold(is_type(stamped_name), |acc, test| {
            java_binary(BinOp::Or, acc, test)
        })
}

fn rewrite_java_record_accessors_stmt(stmt: &mut Statement, binding: &str, type_name: &str) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            rewrite_java_record_accessors_expr(expr, binding, type_name);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_java_record_accessors_expr(init, binding, type_name);
                }
            }
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                rewrite_java_record_accessors_expr(target, binding, type_name);
            }
            rewrite_java_record_accessors_expr(value, binding, type_name);
        }
        StmtKind::Block(body) => {
            for nested in body {
                rewrite_java_record_accessors_stmt(nested, binding, type_name);
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_java_record_accessors_expr(cond, binding, type_name);
            for nested in then_body {
                rewrite_java_record_accessors_stmt(nested, binding, type_name);
            }
            for (elif_cond, elif_body) in elifs {
                rewrite_java_record_accessors_expr(elif_cond, binding, type_name);
                for nested in elif_body {
                    rewrite_java_record_accessors_stmt(nested, binding, type_name);
                }
            }
            if let Some(else_body) = else_body {
                for nested in else_body {
                    rewrite_java_record_accessors_stmt(nested, binding, type_name);
                }
            }
        }
        _ => {}
    }
}

fn rewrite_java_record_accessors_expr(expr: &mut Expression, binding: &str, type_name: &str) {
    if let ExprKind::Call { callee, args, .. } = &mut expr.kind {
        if args.is_empty() {
            if let ExprKind::Member { object, field, .. } = &callee.kind {
                if matches!(object.kind, ExprKind::Ident(ref name) if name == binding)
                    && java_record_has_component(Some(type_name), field)
                {
                    *expr = Expression::new(ExprKind::Member {
                        object: Box::new(Expression::ident(binding)),
                        field: java_record_storage_field(field),
                        null_safe: false,
                    });
                    return;
                }
            }
        }
    }
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            rewrite_java_record_accessors_expr(callee, binding, type_name);
            for arg in args {
                rewrite_java_record_accessors_expr(&mut arg.value, binding, type_name);
            }
        }
        ExprKind::Binary { left, right, .. } => {
            rewrite_java_record_accessors_expr(left, binding, type_name);
            rewrite_java_record_accessors_expr(right, binding, type_name);
        }
        ExprKind::Member { object, .. } => {
            rewrite_java_record_accessors_expr(object, binding, type_name);
        }
        ExprKind::Unary { expr: inner, .. } => {
            rewrite_java_record_accessors_expr(inner, binding, type_name);
        }
        _ => {}
    }
}

fn substitute_java_ident_expr(expr: &mut Expression, name: &str, replacement: &Expression) {
    match &mut expr.kind {
        ExprKind::Ident(ident) if ident == name => {
            *expr = replacement.clone();
        }
        ExprKind::Binary { left, right, .. } => {
            substitute_java_ident_expr(left, name, replacement);
            substitute_java_ident_expr(right, name, replacement);
        }
        ExprKind::Unary { expr: inner, .. }
        | ExprKind::Spread(inner)
        | ExprKind::Await(inner)
        | ExprKind::YieldFrom(inner)
        | ExprKind::Void(inner)
        | ExprKind::Delete(inner)
        | ExprKind::TypeOf(inner)
        | ExprKind::RefLoad(inner) => {
            substitute_java_ident_expr(inner, name, replacement);
        }
        ExprKind::Yield(Some(inner)) => substitute_java_ident_expr(inner, name, replacement),
        ExprKind::Ternary { cond, then, else_ } => {
            substitute_java_ident_expr(cond, name, replacement);
            substitute_java_ident_expr(then, name, replacement);
            substitute_java_ident_expr(else_, name, replacement);
        }
        ExprKind::Member { object, .. } => {
            substitute_java_ident_expr(object, name, replacement);
        }
        ExprKind::Index { object, index, .. } => {
            substitute_java_ident_expr(object, name, replacement);
            substitute_java_ident_expr(index, name, replacement);
        }
        ExprKind::Call { callee, args, .. } => {
            substitute_java_ident_expr(callee, name, replacement);
            for arg in args {
                substitute_java_ident_expr(&mut arg.value, name, replacement);
            }
        }
        ExprKind::Assign { target, value } => {
            substitute_java_ident_expr(target, name, replacement);
            substitute_java_ident_expr(value, name, replacement);
        }
        ExprKind::Array(elems) => {
            for elem in elems {
                substitute_java_ident_expr(&mut elem.value, name, replacement);
            }
        }
        ExprKind::Tuple(elems) | ExprKind::Set(elems) | ExprKind::Sequence(elems) => {
            for elem in elems {
                substitute_java_ident_expr(elem, name, replacement);
            }
        }
        ExprKind::New { class, args } => {
            substitute_java_ident_expr(class, name, replacement);
            for arg in args {
                substitute_java_ident_expr(&mut arg.value, name, replacement);
            }
        }
        ExprKind::Lambda { .. } => {}
        _ => {}
    }
}

fn java_or_exprs(mut exprs: Vec<Expression>) -> Option<Expression> {
    let first = exprs.pop()?;
    Some(
        exprs
            .into_iter()
            .fold(first, |acc, expr| java_binary(BinOp::Or, expr, acc)),
    )
}

fn java_var_decl(name: &str, init: Option<Expression>) -> Statement {
    java_var_decl_typed(name, None, init)
}

fn java_var_decl_typed(
    name: &str,
    type_hint: Option<String>,
    init: Option<Expression>,
) -> Statement {
    Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(name.to_string()),
            type_hint,
            init,
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    })
}

fn java_assign_stmt(name: &str, value: Expression) -> Statement {
    Statement::new(StmtKind::Assign {
        targets: vec![Expression::ident(name)],
        value,
    })
}

fn walk_switch_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner();
    let subject = walk_expr_inner(&mut inner)?;
    let mut arms: Vec<(Vec<JavaSwitchLabel>, Expression)> = Vec::new();
    let mut default_expr: Option<Expression> = None;

    for arm in inner {
        if arm.as_rule() != Rule::switch_expr_arm {
            continue;
        }
        let mut labels = Vec::new();
        let mut value = None;
        let mut is_default = false;
        let mut ai = arm.into_inner().peekable();
        let src = ai
            .peek()
            .map(|p| p.as_str().to_string())
            .unwrap_or_default();
        if src.trim() == "default" {
            is_default = true;
            ai.next();
        }
        for p in ai {
            match p.as_rule() {
                Rule::switch_label => {
                    labels.push(java_switch_label_expr(walk_switch_label(p)?));
                }
                Rule::switch_rule_body => {
                    value = java_switch_rule_body_expr(p)?;
                }
                _ => {}
            }
        }
        if let Some(expr) = value {
            if is_default || labels.is_empty() {
                default_expr = Some(expr);
            } else {
                arms.push((labels, expr));
            }
        }
    }

    let mut result = default_expr.unwrap_or_else(Expression::null);
    let subject_name = "__java_switch_expr_subject";
    for (labels, value) in arms.into_iter().rev() {
        let cond = java_or_exprs(
            labels
                .iter()
                .map(|label| java_switch_label_match(subject_name, label))
                .collect(),
        )
        .unwrap_or_else(|| Expression::bool(false));
        let mut value = value;
        for label in &labels {
            if let JavaSwitchLabel::Pattern { binding, .. } = label {
                substitute_java_ident_expr(&mut value, binding, &Expression::ident(subject_name));
            }
        }
        result = java_ternary(cond, value, result);
    }
    Ok(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: vec![Param {
                name: subject_name.to_string(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            }],
            body: LambdaBody::Expr(Box::new(result)),
            captures: vec![],
            is_async: false,
        })),
        args: vec![Argument::positional(subject)],
        optional: false,
    }))
}

fn java_switch_rule_body_expr(pair: Pair<Rule>) -> Result<Option<Expression>, String> {
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::expression => return Ok(Some(walk_expression(p)?)),
            Rule::expression_statement => {
                if let Some(expr_p) = p.into_inner().next() {
                    return Ok(Some(walk_expression(expr_p)?));
                }
            }
            Rule::block_statement => {
                for stmt_pair in p.into_inner() {
                    if stmt_pair.as_rule() == Rule::yield_statement {
                        let expr = stmt_pair
                            .into_inner()
                            .find(|inner| !is_kw(inner.as_rule()))
                            .map(walk_expression)
                            .transpose()?;
                        return Ok(expr);
                    }
                }
            }
            _ => {}
        }
    }
    Ok(None)
}

fn walk_try(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut body: Vec<Statement> = Vec::new();
    let mut catches: Vec<CatchClause> = Vec::new();
    let mut finally: Option<Vec<Statement>> = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::resource => {}
            Rule::block_statement | Rule::function_body_block => {
                body = walk_block(p)?;
            }
            Rule::catch_clause => {
                let mut ci = p.into_inner().peekable();
                if ci.peek().map(|x| x.as_rule()) == Some(Rule::final_kw) {
                    ci.next();
                }
                let mut types: Vec<String> = Vec::new();
                while ci.peek().map(|x| x.as_rule()) == Some(Rule::type_ref) {
                    let ty = extract_ref_name(&ci.next().unwrap());
                    let ty = if ty.starts_with("java.") {
                        java_type_simple_name(&ty).to_string()
                    } else {
                        ty
                    };
                    types.push(ty);
                }
                let var_name = ci.next().map(|p| p.as_str().to_string());
                let catch_body = ci
                    .next()
                    .map(|b| walk_block(b))
                    .transpose()?
                    .unwrap_or_default();
                if types.is_empty() {
                    types.push("Exception".to_string());
                }
                catches.push(CatchClause {
                    types,
                    var_name,
                    stack_var: None,
                    body: catch_body,
                    when_clause: None,
                });
            }
            Rule::finally_clause => {
                if let Some(blk) = p.into_inner().next() {
                    finally = Some(walk_block(blk)?);
                }
            }
            _ => {}
        }
    }

    Ok(StmtKind::Try {
        body,
        catches,
        else_body: None,
        finally,
    })
}

// ════════════════════════════════════════════════════════════════════════════
// Parameters
// ════════════════════════════════════════════════════════════════════════════

fn walk_params(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    let mut out = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::param {
            out.push(walk_param(p)?);
        }
    }
    Ok(out)
}

fn walk_param(pair: Pair<Rule>) -> Result<Param, String> {
    let mut inner = pair.into_inner().peekable();

    // annotations
    while inner.peek().map(|p| p.as_rule()) == Some(Rule::annotation) {
        inner.next();
    }
    // optional final
    if inner.peek().map(|p| p.as_rule()) == Some(Rule::final_kw) {
        inner.next();
    }
    // type_ref
    let type_hint = if inner.peek().map(|p| p.as_rule()) == Some(Rule::type_ref) {
        Some(extract_ref_name(&inner.next().unwrap()))
    } else {
        None
    };
    // varargs
    let is_rest = if inner.peek().map(|p| p.as_rule()) == Some(Rule::varargs_marker) {
        inner.next();
        true
    } else {
        false
    };
    let name = inner
        .next()
        .ok_or("param: missing name")?
        .as_str()
        .to_string();
    // skip dim_suffix(s)
    while inner.peek().map(|p| p.as_rule()) == Some(Rule::dim_suffix) {
        inner.next();
    }

    Ok(Param {
        name,
        type_hint,
        default: None,
        pass_by: PassBy::Value,
        is_rest,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    })
}

// ════════════════════════════════════════════════════════════════════════════
// Expressions
// ════════════════════════════════════════════════════════════════════════════

fn walk_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    match pair.as_rule() {
        Rule::expression => {
            let mut parts: Vec<Expression> = pair
                .into_inner()
                .map(walk_expression)
                .collect::<Result<_, _>>()?;
            if parts.len() == 1 {
                Ok(parts.remove(0))
            } else {
                Ok(Expression::new(ExprKind::Sequence(parts)))
            }
        }
        Rule::assignment_expression => walk_assignment(pair),
        Rule::ternary_expression => walk_ternary(pair),
        Rule::binop_expression => walk_binop(pair),
        Rule::instanceof_expression => walk_instanceof(pair),
        Rule::unary_expression => walk_unary(pair),
        Rule::cast_expression => {
            let mut ci = pair.into_inner();
            let cast_type = ci.next(); // cast_type
            if let (Some(cast_type), Some(operand)) = (cast_type, ci.next()) {
                let expr = walk_expression(operand)?;
                let ty = cast_type.as_str();
                let simple_ty = java_type_simple_name(ty);
                if matches!(simple_ty, "int" | "Integer") {
                    Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__j_i32")),
                        args: vec![Argument::positional(Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident("__java_trunc_cast")),
                            args: vec![Argument::positional(expr)],
                            optional: false,
                        }))],
                        optional: false,
                    }))
                } else if matches!(simple_ty, "double" | "Double" | "float" | "Float") {
                    Ok(Expression::new(ExprKind::Cast {
                        type_name: simple_ty.to_string(),
                        expr: Box::new(expr),
                    }))
                } else if simple_ty == "char" {
                    Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__j_from_char_code")),
                        args: vec![Argument::positional(java_char_numeric_cast_expr(expr))],
                        optional: false,
                    }))
                } else if let Some(callee) = java_numeric_cast_fn(ty) {
                    Ok(Expression::new(ExprKind::Cast {
                        type_name: simple_ty.to_string(),
                        expr: Box::new(Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident(callee)),
                            args: vec![Argument::positional(expr)],
                            optional: false,
                        })),
                    }))
                } else {
                    Ok(expr)
                }
            } else {
                Ok(Expression::null())
            }
        }
        Rule::postfix_expression => walk_postfix(pair),
        Rule::primary_chain => walk_primary_chain(pair),
        Rule::primary_atom => walk_primary_atom(pair),
        Rule::lambda_expression => walk_lambda(pair),
        Rule::switch_expression => walk_switch_expression(pair),
        _ => {
            let mut inner = pair.into_inner();
            if let Some(first) = inner.next() {
                walk_expression(first)
            } else {
                Ok(Expression::null())
            }
        }
    }
}

fn walk_assignment(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner().peekable();
    let first = inner.next().ok_or("assignment: empty")?;

    // Check if next is an assignment_op
    if inner.peek().map(|p| p.as_rule()) == Some(Rule::assignment_op) {
        let op_str = inner.next().unwrap().as_str().to_string();
        let rhs = walk_expression(inner.next().ok_or("assignment: missing rhs")?)?;
        let lhs = walk_expression(first)?;

        if op_str == "=" {
            return Ok(Expression::new(ExprKind::Assign {
                target: Box::new(lhs),
                value: Box::new(rhs),
            }));
        }
        // Compound assignment: `x += v` → `x = x + v`
        let bin_op = compound_op_to_binop(&op_str);
        return Ok(Expression::new(ExprKind::Assign {
            target: Box::new(lhs.clone()),
            value: Box::new(Expression::new(ExprKind::Binary {
                op: bin_op,
                left: Box::new(lhs),
                right: Box::new(rhs),
            })),
        }));
    }

    walk_expression(first)
}

fn walk_ternary(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner();
    let cond = walk_expression(inner.next().ok_or("ternary: missing cond")?)?;
    if let Some(then_p) = inner.next() {
        let then_e = walk_expression(then_p)?;
        let else_e = walk_expression(inner.next().ok_or("ternary: missing else")?)?;
        Ok(Expression::new(ExprKind::Ternary {
            cond: Box::new(cond),
            then: Box::new(then_e),
            else_: Box::new(else_e),
        }))
    } else {
        Ok(cond)
    }
}

fn walk_binop(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner();
    let first = walk_expression(inner.next().ok_or("binop: missing lhs")?)?;
    let mut operands = vec![first];
    let mut ops = Vec::new();

    while let Some(op_pair) = inner.next() {
        let rhs = walk_expression(inner.next().ok_or("binop: missing rhs")?)?;
        ops.push(str_to_binop(op_pair.as_str().trim()));
        operands.push(rhs);
    }
    Ok(build_java_binop_precedence(operands, ops))
}

fn build_java_binop_precedence(mut operands: Vec<Expression>, mut ops: Vec<BinOp>) -> Expression {
    for level in [
        &[BinOp::Mul, BinOp::Div, BinOp::Mod][..],
        &[BinOp::Add, BinOp::Sub][..],
        &[BinOp::Shl, BinOp::Shr, BinOp::UShr][..],
        &[BinOp::Lt, BinOp::LtEq, BinOp::Gt, BinOp::GtEq][..],
        &[BinOp::Eq, BinOp::NotEq][..],
        &[BinOp::BitAnd][..],
        &[BinOp::BitXor][..],
        &[BinOp::BitOr][..],
        &[BinOp::And][..],
        &[BinOp::Or][..],
    ] {
        let mut i = 0;
        while i < ops.len() {
            if level.contains(&ops[i]) {
                let op = ops.remove(i);
                let left = operands.remove(i);
                let right = operands.remove(i);
                operands.insert(i, java_binary_with_string_concat(op, left, right));
            } else {
                i += 1;
            }
        }
    }
    operands.into_iter().next().unwrap_or_else(Expression::null)
}

fn java_binary_with_string_concat(op: BinOp, left: Expression, right: Expression) -> Expression {
    if op == BinOp::Add
        && (is_java_string_concat_operand(&left) || is_java_string_concat_operand(&right))
    {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__java_string_concat")),
            args: vec![Argument::positional(left), Argument::positional(right)],
            optional: false,
        })
    } else {
        let effective_op = if op == BinOp::Div
            && !is_java_double_arithmetic_expr(&left)
            && !is_java_double_arithmetic_expr(&right)
        {
            BinOp::IDiv
        } else {
            op
        };
        let expr = Expression::new(ExprKind::Binary {
            op: effective_op,
            left: Box::new(left),
            right: Box::new(right),
        });
        if matches!(effective_op, BinOp::Add | BinOp::Sub | BinOp::Mul)
            && contains_java_integer_bound_constant(&expr)
        {
            Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__j_i32")),
                args: vec![Argument::positional(expr)],
                optional: false,
            })
        } else {
            expr
        }
    }
}

fn is_java_string_concat_operand(expr: &Expression) -> bool {
    match expr.kind {
        ExprKind::Lit(Literal::Str(_)) => true,
        ExprKind::Call { ref callee, .. } => {
            matches!(callee.kind, ExprKind::Ident(ref name) if name == "__java_string_concat")
        }
        _ => false,
    }
}

fn contains_java_integer_bound_constant(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Member { object, field, .. } => {
            matches!(&object.kind, ExprKind::Ident(name) if name == "Integer")
                && matches!(field.as_str(), "MAX_VALUE" | "MIN_VALUE")
        }
        ExprKind::Unary { expr, .. } => contains_java_integer_bound_constant(expr),
        ExprKind::Binary { left, right, .. } => {
            contains_java_integer_bound_constant(left)
                || contains_java_integer_bound_constant(right)
        }
        ExprKind::Call { args, .. } => args
            .iter()
            .any(|arg| contains_java_integer_bound_constant(&arg.value)),
        _ => false,
    }
}

fn is_java_double_arithmetic_expr(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Lit(Literal::Float(_)) => true,
        ExprKind::Ident(name) => {
            JAVA_DOUBLE_VARS.with(|vars| vars.borrow().contains(name.as_str()))
        }
        ExprKind::Member { object, field, .. } => {
            matches!(
                java_expr_dotted_name(object).as_deref(),
                Some("Double") | Some("java.lang.Double") | Some("Long") | Some("java.lang.Long")
            ) && matches!(field.as_str(), "MAX_VALUE" | "MIN_VALUE")
        }
        ExprKind::Call { callee, .. } => matches!(
            &callee.kind,
            ExprKind::Ident(name)
                if matches!(
                    name.as_str(),
                    "Double.parseDouble" | "Double.valueOf" | "Float.parseFloat" | "Float.valueOf"
                )
        ),
        ExprKind::Cast { type_name, .. } => {
            matches!(
                java_type_simple_name(type_name),
                "double" | "Double" | "float" | "Float"
            )
        }
        ExprKind::Binary {
            op, left, right, ..
        } if matches!(
            op,
            BinOp::Add | BinOp::Sub | BinOp::Mul | BinOp::Div | BinOp::IDiv | BinOp::Mod
        ) =>
        {
            is_java_double_arithmetic_expr(left) || is_java_double_arithmetic_expr(right)
        }
        ExprKind::Unary { expr, .. } => is_java_double_arithmetic_expr(expr),
        _ => false,
    }
}

fn java_expr_is_long_value(expr: &Expression) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Cast { type_name, .. } if matches!(java_type_simple_name(type_name), "long" | "Long")
    )
}

fn java_print_arg(arg: Argument) -> Argument {
    if let Some(value) = java_wrapper_constant_print_string(&arg.value) {
        return Argument::positional(value);
    }
    if is_java_double_arithmetic_expr(&arg.value) {
        Argument::positional(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__java_double_to_string")),
            args: vec![Argument::positional(arg.value)],
            optional: false,
        }))
    } else {
        arg
    }
}

fn java_wrapper_constant_print_string(expr: &Expression) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &expr.kind else {
        return None;
    };
    let type_name = java_expr_dotted_name(object)?;
    let text = match (type_name.as_str(), field.as_str()) {
        ("Long", "MAX_VALUE") | ("java.lang.Long", "MAX_VALUE") => "9.223372036854776E18",
        ("Long", "MIN_VALUE") | ("java.lang.Long", "MIN_VALUE") => "-9.223372036854776E18",
        ("Double", "MAX_VALUE") | ("java.lang.Double", "MAX_VALUE") => "1.7976931348623157E308",
        _ => return None,
    };
    Some(Expression::string(text))
}

fn java_numeric_cast_fn(ty: &str) -> Option<&'static str> {
    match ty {
        "byte" | "Byte" => Some("__j_byte"),
        "short" | "Short" => Some("__j_short"),
        "int" | "long" | "Integer" | "Long" => Some("__java_trunc_cast"),
        _ => None,
    }
}

fn java_numeric_width_fn(ty: &str) -> Option<&'static str> {
    match java_type_simple_name(ty) {
        "byte" | "Byte" => Some("__j_byte"),
        "short" | "Short" => Some("__j_short"),
        _ => None,
    }
}

fn java_char_numeric_cast_expr(expr: Expression) -> Expression {
    match expr.kind {
        ExprKind::Lit(Literal::Char(_)) => Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__java_char_ord")),
            args: vec![Argument::positional(expr)],
            optional: false,
        }),
        ExprKind::Lit(Literal::Str(ref value)) if value.chars().count() == 1 => {
            Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__java_char_ord")),
                args: vec![Argument::positional(expr)],
                optional: false,
            })
        }
        ExprKind::Binary { op, left, right } => Expression::new(ExprKind::Binary {
            op,
            left: Box::new(java_char_numeric_cast_expr(*left)),
            right: Box::new(java_char_numeric_cast_expr(*right)),
        }),
        ExprKind::Index {
            object,
            index,
            null_safe,
        } => Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__java_char_ord")),
            args: vec![Argument::positional(Expression::new(ExprKind::Index {
                object,
                index,
                null_safe,
            }))],
            optional: false,
        }),
        ExprKind::Call { callee, args, .. }
            if matches!(&callee.kind, ExprKind::Ident(name) if name == "__java_string_concat")
                && args.len() == 2 =>
        {
            let mut args = args.into_iter();
            let left = args.next().expect("left").value;
            let right = args.next().expect("right").value;
            Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(java_char_numeric_cast_expr(left)),
                right: Box::new(java_char_numeric_cast_expr(right)),
            })
        }
        ExprKind::Unary { op, expr } => Expression::new(ExprKind::Unary {
            op,
            expr: Box::new(java_char_numeric_cast_expr(*expr)),
        }),
        _ => expr,
    }
}

fn java_expr_is_char_numeric_source(
    expr: &Expression,
    local_types: &HashMap<String, String>,
) -> bool {
    match &expr.kind {
        ExprKind::Lit(Literal::Char(_)) => true,
        ExprKind::Ident(name) => local_types
            .get(name)
            .is_some_and(|ty| java_type_simple_name(ty) == "char"),
        ExprKind::Index { .. } => true,
        ExprKind::Call { callee, .. } => matches!(
            &callee.kind,
            ExprKind::Member { field, .. } if field == "charAt"
        ),
        ExprKind::Cast { type_name, .. } => java_type_simple_name(type_name) == "char",
        _ => false,
    }
}

fn java_expr_is_string_value(expr: &Expression, local_types: &HashMap<String, String>) -> bool {
    match &expr.kind {
        ExprKind::Lit(Literal::Str(value)) => value.chars().count() != 1,
        ExprKind::Ident(name) => {
            local_types
                .get(name)
                .is_some_and(|ty| java_type_simple_name(ty) == "String")
                || JAVA_STRING_VARS.with(|vars| vars.borrow().contains(name.as_str()))
        }
        ExprKind::Call { callee, .. } => matches!(
            &callee.kind,
            ExprKind::Ident(name) if name == "__java_string_concat"
        ),
        ExprKind::Cast { type_name, .. } => java_type_simple_name(type_name) == "String",
        _ => false,
    }
}

fn java_cast_char_numeric_operand(
    expr: Expression,
    local_types: &HashMap<String, String>,
) -> Expression {
    if java_expr_is_char_numeric_source(&expr, local_types) {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__java_char_ord")),
            args: vec![Argument::positional(expr)],
            optional: false,
        })
    } else {
        java_char_numeric_cast_expr(expr)
    }
}

fn walk_instanceof(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner();
    let mut base = walk_expression(inner.next().ok_or("instanceof: empty")?)?;

    let mut saw_instanceof = false;
    // The value being tested by the most recent `instanceof`, kept so a
    // trailing pattern-binding ident can capture it.
    let mut pending: Option<(String, Expression)> = None;
    for p in inner {
        match p.as_rule() {
            Rule::instanceof_kw => {
                saw_instanceof = true;
            }
            Rule::type_ref if saw_instanceof => {
                let type_name = extract_ref_name(&p);
                let subject = base.clone();
                base = java_instanceof_match_expr(&subject, &type_name);
                pending = Some((type_name, subject));
                saw_instanceof = false;
            }
            Rule::ident_name => {
                // `instanceof T var` (JLS §14.30.1) — record the flow-scoped
                // binding for walk_if to inject into the then-body.
                if let Some((type_name, subject)) = pending.take() {
                    JAVA_INSTANCEOF_BINDINGS.with(|b| {
                        b.borrow_mut()
                            .push((p.as_str().to_string(), type_name, subject))
                    });
                }
            }
            _ => {}
        }
    }
    Ok(base)
}

/// The type-test expression for `subject instanceof Type`. Identical to the
/// switch-pattern path when the subject is a plain name, so enum membership is
/// answered the same way in both; otherwise the shared type test.
fn java_instanceof_match_expr(subject: &Expression, type_name: &str) -> Expression {
    if let ExprKind::Ident(name) = &subject.kind {
        return java_pattern_type_match_expr(name, type_name);
    }
    java_type_test_expr(subject, type_name)
}

fn walk_unary(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner();
    let first = inner.next().ok_or("unary: empty")?;

    if first.as_rule() == Rule::unary_op {
        let op_str = first.as_str();
        let operand = walk_expression(inner.next().ok_or("unary: missing operand")?)?;
        if op_str == "--" && matches!(operand.kind, ExprKind::Lit(_)) {
            return Ok(operand);
        }
        let op = match op_str {
            "++" => UnaryOp::PreInc,
            "--" => UnaryOp::PreDec,
            "!" => UnaryOp::Not,
            "-" => UnaryOp::Neg,
            "+" => UnaryOp::Pos,
            "~" => UnaryOp::BitNot,
            _ => UnaryOp::Not,
        };
        return Ok(Expression::new(ExprKind::Unary {
            op,
            expr: Box::new(operand),
        }));
    }
    walk_expression(first)
}

fn walk_postfix(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner();
    let base = walk_expression(inner.next().ok_or("postfix: empty")?)?;
    if let Some(op) = inner.next() {
        let unop = match op.as_str() {
            "++" => UnaryOp::PostInc,
            "--" => UnaryOp::PostDec,
            _ => UnaryOp::PostInc,
        };
        Ok(Expression::new(ExprKind::Unary {
            op: unop,
            expr: Box::new(base),
        }))
    } else {
        Ok(base)
    }
}

fn java_inner_new_receiver_type(receiver: &Expression) -> Option<String> {
    match &receiver.kind {
        ExprKind::Ident(name) => JAVA_LOCAL_TYPES.with(|types| types.borrow().get(name).cloned()),
        ExprKind::New { class, .. } => java_expr_dotted_name(class),
        ExprKind::Call { callee, .. } => {
            if let ExprKind::Member { object, field, .. } = &callee.kind {
                let receiver_type = java_inner_new_receiver_type(object)?;
                java_class_method_return_type(&receiver_type, field)
            } else {
                None
            }
        }
        ExprKind::Member { object, field, .. } => {
            let owner = java_inner_new_receiver_type(object)?;
            java_class_field_type(&owner, field)
        }
        _ => None,
    }
}

fn walk_primary_chain(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner();
    let mut current = walk_expression(inner.next().ok_or("chain: empty")?)?;

    for chain in inner {
        // chain_suffix is a wrapper rule — unwrap to get the inner rule
        let chain = if chain.as_rule() == Rule::chain_suffix {
            chain.into_inner().next().unwrap_or_else(|| unreachable!())
        } else {
            chain
        };
        match chain.as_rule() {
            Rule::inner_new_suffix => {
                let mut ci = chain.into_inner().peekable();
                let class_name = ci
                    .next()
                    .map(|p| extract_ref_name(&p))
                    .unwrap_or_else(|| "Object".to_string());
                if ci.peek().map(|x| x.as_rule()) == Some(Rule::type_args) {
                    ci.next();
                }
                let mut args = if let Some(al) = ci.next() {
                    if al.as_rule() == Rule::argument_list {
                        walk_arguments(al)?
                    } else {
                        vec![]
                    }
                } else {
                    vec![]
                };
                let owner_type = java_inner_new_receiver_type(&current);
                let qualified = owner_type
                    .as_deref()
                    .map(|owner| {
                        if class_name.contains('.') {
                            class_name.clone()
                        } else {
                            format!("{owner}.{class_name}")
                        }
                    })
                    .unwrap_or_else(|| class_name.clone());
                args.insert(0, Argument::positional(current));
                current = Expression::new(ExprKind::New {
                    class: Box::new(Expression::ident(&qualified)),
                    args,
                });
            }
            Rule::method_call_suffix => {
                let mut ci = chain.into_inner().peekable();
                if ci.peek().map(|x| x.as_rule()) == Some(Rule::type_args) {
                    ci.next();
                }
                let method_name = ci
                    .next()
                    .ok_or("method call: missing name")?
                    .as_str()
                    .to_string();
                if ci.peek().map(|x| x.as_rule()) == Some(Rule::type_args) {
                    ci.next();
                }
                let args = if let Some(al) = ci.next() {
                    if al.as_rule() == Rule::argument_list {
                        walk_arguments(al)?
                    } else {
                        vec![]
                    }
                } else {
                    vec![]
                };
                current = normalise_method_call(current, method_name, args);
            }
            Rule::member_access_suffix => {
                let field = chain
                    .into_inner()
                    .next()
                    .ok_or("member: empty")?
                    .as_str()
                    .to_string();
                if let Some(constant) = java_double_constant_expr(&current, &field) {
                    current = constant;
                } else if let Some(constant) = java_locale_constant_expr(&current, &field) {
                    current = constant;
                } else if let Some(constant) = java_calendar_constant_expr(&current, &field) {
                    current = constant;
                } else {
                    current = Expression::new(ExprKind::Member {
                        object: Box::new(current),
                        field,
                        null_safe: false,
                    });
                }
            }
            Rule::index_suffix => {
                let idx = walk_expression(chain.into_inner().next().ok_or("index: empty")?)?;
                current = Expression::new(ExprKind::Index {
                    object: Box::new(current),
                    index: Box::new(idx),
                    null_safe: false,
                });
            }
            Rule::call_suffix => {
                // Bare function call: callee(args) — the base is the callee.
                let mut ci = chain.into_inner().peekable();
                if ci.peek().map(|x| x.as_rule()) == Some(Rule::type_args) {
                    ci.next();
                }
                let args = if let Some(al) = ci.next() {
                    if al.as_rule() == Rule::argument_list {
                        walk_arguments(al)?
                    } else {
                        vec![]
                    }
                } else {
                    vec![]
                };
                if let ExprKind::Member {
                    object,
                    field,
                    null_safe: _,
                } = current.kind
                {
                    current = normalise_method_call(*object, field, args);
                    continue;
                }
                if let ExprKind::Ident(name) = &current.kind {
                    if let Some(callee) = java_dotted_static_call_name(name) {
                        current = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident(callee)),
                            args,
                            optional: false,
                        });
                        continue;
                    }
                    if matches!(name.as_str(), "wait" | "notify" | "notifyAll") {
                        current = normalise_method_call(
                            Expression::new(ExprKind::This),
                            name.clone(),
                            args,
                        );
                        continue;
                    }
                }
                current = Expression::new(ExprKind::Call {
                    callee: Box::new(current),
                    args,
                    optional: false,
                });
            }
            _ => {}
        }
    }
    Ok(current)
}

fn java_dotted_static_call_name(name: &str) -> Option<&'static str> {
    Some(match name {
        "Executors.newFixedThreadPool" | "java.util.concurrent.Executors.newFixedThreadPool" => {
            "__j_exec_new"
        }
        _ => return None,
    })
}

fn java_double_constant_expr(object: &Expression, field: &str) -> Option<Expression> {
    let type_name = java_expr_dotted_name(object)?;
    let f64_lit = |value: f64| Expression::new(ExprKind::Lit(Literal::Float(value)));
    let div = |left: Expression, right: Expression| {
        Expression::new(ExprKind::Binary {
            op: BinOp::Div,
            left: Box::new(left),
            right: Box::new(right),
        })
    };
    if type_name == "StrictMath" || type_name == "java.lang.StrictMath" {
        return match field {
            "PI" => Some(f64_lit(std::f64::consts::PI)),
            "E" => Some(f64_lit(std::f64::consts::E)),
            _ => None,
        };
    }
    if type_name != "Double" && type_name != "java.lang.Double" {
        return None;
    }
    match field {
        "MAX_VALUE" => Some(f64_lit(f64::MAX)),
        "MIN_VALUE" => Some(f64_lit(f64::from_bits(1))),
        "MIN_EXPONENT" => Some(Expression::int(-1022)),
        "NaN" => Some(div(f64_lit(0.0), f64_lit(0.0))),
        "POSITIVE_INFINITY" => Some(div(f64_lit(1.0), f64_lit(0.0))),
        "NEGATIVE_INFINITY" => Some(div(f64_lit(-1.0), f64_lit(0.0))),
        _ => None,
    }
}

fn java_locale_constant_expr(object: &Expression, field: &str) -> Option<Expression> {
    let owner = java_expr_dotted_name(object)?;
    if !matches!(owner.as_str(), "Locale" | "java.util.Locale") {
        return None;
    }
    let value = match field {
        "FRANCE" => "FR",
        "GERMANY" => "DE",
        "ITALY" => "IT",
        "US" => "US",
        "UK" => "UK",
        "JAPAN" => "JP",
        "CANADA" => "CA",
        "CANADA_FRENCH" => "FR_CA",
        _ => return None,
    };
    Some(Expression::string(value))
}

fn java_calendar_constant_expr(object: &Expression, field: &str) -> Option<Expression> {
    let owner = java_expr_dotted_name(object)?;
    if !matches!(owner.as_str(), "Calendar" | "java.util.Calendar") {
        return None;
    }
    match field {
        "MILLISECOND" => Some(Expression::int(14)),
        _ => None,
    }
}

/// Build one PrintStream write. Basic print/println are profile/common
/// emitters; append/format still use the legacy `__j_*` runtime until those
/// composite formatting pieces are migrated too.
fn java_print_stream_write(method: &str, args: Vec<Argument>) -> Expression {
    let build = |name: &str, args: Vec<Argument>| {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident(name)),
            args,
            optional: false,
        })
    };
    let first_or_empty = |args: Vec<Argument>| {
        args.into_iter()
            .next()
            .unwrap_or_else(|| Argument::positional(Expression::string("")))
    };
    match method {
        "println" => build("__java_println", vec![java_print_arg(first_or_empty(args))]),
        "append" if args.len() == 3 => {
            // append(csq, start, end) → write csq.substring(start, end)
            let mut it = args.into_iter();
            let csq = it.next().expect("csq").value;
            let start = it.next().expect("start");
            let end = it.next().expect("end");
            let sub = Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(csq),
                    field: "substring".to_string(),
                    null_safe: false,
                })),
                args: vec![start, end],
                optional: false,
            });
            build("__java_print", vec![Argument::positional(sub)])
        }
        "print" => build("__java_print", vec![java_print_arg(first_or_empty(args))]),
        "append" => build("__java_print", vec![java_print_arg(first_or_empty(args))]),
        // printf | format
        _ => {
            let mut it = args.into_iter();
            let mut fmt = it
                .next()
                .unwrap_or_else(|| Argument::positional(Expression::string("")));
            let mut rest: Vec<ArrayElement> = it
                .map(|arg| ArrayElement {
                    key: None,
                    value: arg.value,
                    spread: false,
                    by_ref: false,
                })
                .collect();
            java_rewrite_printstream_format_literal(&mut fmt.value, &mut rest);
            build(
                "__java_printf",
                vec![
                    fmt,
                    Argument::positional(Expression::new(ExprKind::Array(rest))),
                ],
            )
        }
    }
}

fn java_rewrite_printstream_format_literal(fmt: &mut Expression, args: &mut [ArrayElement]) {
    let ExprKind::Lit(Literal::Str(text)) = &mut fmt.kind else {
        return;
    };

    if text == "%,d" {
        *text = "%s".to_string();
        if let Some(first) = args.first_mut() {
            first.value = java_call(
                Expression::ident("__java_format_grouped_int"),
                vec![first.value.clone()],
            );
        }
        return;
    }
    if text == "%e" || text == "%E" {
        let helper = if text == "%E" {
            "__java_format_exp_upper"
        } else {
            "__java_format_exp_lower"
        };
        *text = "%s".to_string();
        if let Some(first) = args.first_mut() {
            first.value = java_call(Expression::ident(helper), vec![first.value.clone()]);
        }
        return;
    }

    *text = text
        .replace("%n", "\n")
        .replace("%b", "%s")
        .replace("%g", "%.5f");
}

/// Normalise Java-specific call patterns to a compiler-friendly shape.
fn normalise_method_call(receiver: Expression, method: String, args: Vec<Argument>) -> Expression {
    let receiver = java_reflection_indexed_token(&receiver).unwrap_or(receiver);

    if let Some(interface_name) = java_interface_super_receiver(&receiver) {
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(Expression::new(ExprKind::This)),
                field: java_interface_default_method_name(&interface_name, &method),
                null_safe: false,
            })),
            args,
            optional: false,
        });
    }

    if args.is_empty() {
        if method == "getName" {
            if let Some(name) = java_reflection_token_name(&receiver) {
                return Expression::string(&name);
            }
        }
        if method == "getParameterCount" {
            if let Some(count) = java_reflection_token_param_count(&receiver) {
                return Expression::int(count as i64);
            }
        }
        if method == "getType" {
            if let Some(type_name) = java_reflection_token_type_name(&receiver) {
                return Expression::string(&type_name);
            }
        }
        if method == "getReturnType" {
            if let Some(type_name) = java_reflection_token_return_type_name(&receiver) {
                return Expression::string(&type_name);
            }
        }
        if method == "getParameterTypes" {
            if let Some(types) = java_reflection_token_param_types(&receiver) {
                return java_string_array_expr(types);
            }
        }
        if method == "getModifiers" {
            if let Some(modifiers) = java_reflection_token_modifiers(&receiver) {
                return Expression::int(modifiers);
            }
        }
        if method == "setAccessible" {
            if java_reflection_token_kind(&receiver).is_some() {
                return Expression::null();
            }
        }
    }

    if let Some(expr) = java_reflection_token_operation(&receiver, &method, &args) {
        return expr;
    }

    if args.len() == 1
        && java_member_chain_ends_with(&receiver, "Modifier")
        && matches!(
            method.as_str(),
            "isPublic" | "isPrivate" | "isProtected" | "isStatic" | "isFinal" | "isAbstract"
        )
    {
        if let Some(expr) = java_modifier_static_predicate(&method, &args[0].value) {
            return expr;
        }
    }

    if method == "isInstance" && args.len() == 1 {
        if let ExprKind::Lit(Literal::Str(type_name)) = &receiver.kind {
            // Every type — user-declared or stdlib — answers through the SHARED
            // type test (`ExprKind::IsType`). This used to fork on a hardcoded
            // ~30-name list into `__java_class_is_instance`, an entire
            // hand-written bytecode ladder living in the Java emitter.
            // `X.class` is a string literal naming the type, so the simple
            // name is all there is to match on.
            return java_type_test_expr(&args[0].value, java_type_simple_name(type_name));
        }
    }

    // PrintStream writes (JLS java.io.PrintStream): `System.out`, a
    // PrintStream-typed local, or a chained write (`….append(x).format(…)`).
    // Basic print/println route through profile emitters; formatter/append
    // still route through the legacy runtime.
    if matches!(
        method.as_str(),
        "println" | "print" | "printf" | "format" | "append"
    ) {
        let stream_receiver = match &receiver.kind {
            ExprKind::Member { object, field, .. } => {
                matches!(&object.kind, ExprKind::Ident(n) if n == "System") && field == "out"
            }
            ExprKind::Ident(name) => {
                name == "__j_out"
                    || name == "__java_out"
                    || JAVA_PRINTSTREAM_VARS.with(|vars| vars.borrow().contains(name.as_str()))
            }
            ExprKind::Call { callee, .. } => matches!(
                &callee.kind,
                ExprKind::Ident(n) if matches!(n.as_str(), "__j_print" | "__j_println" | "__j_printf" | "__java_print" | "__java_println" | "__java_printf")
            ),
            ExprKind::Sequence(items) => matches!(
                items.last().map(|e| &e.kind),
                Some(ExprKind::Call { callee, .. }) if matches!(
                    &callee.kind,
                    ExprKind::Ident(n) if matches!(n.as_str(), "__j_print" | "__j_println" | "__j_printf" | "__java_print" | "__java_println" | "__java_printf")
                )
            ),
            _ => false,
        };
        if stream_receiver {
            let write = java_print_stream_write(&method, args);
            // Chained receivers carry earlier writes as side effects —
            // keep them, in order, ahead of this one.
            return match receiver.kind {
                ExprKind::Member { .. } | ExprKind::Ident(_) => write,
                _ => Expression::new(ExprKind::Sequence(vec![receiver, write])),
            };
        }
    }

    if method == "run"
        && args.is_empty()
        && matches!(&receiver.kind, ExprKind::Ident(n) if JAVA_RUNNABLE_VARS.with(|vars| vars.borrow().contains(n.as_str())))
    {
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__j_runnable_run")),
            args: vec![Argument::positional(receiver)],
            optional: false,
        });
    }

    if method == "filter" && args.len() == 1 && java_optional_receiver(&receiver) {
        let mut call_args = vec![Argument::positional(receiver)];
        call_args.extend(args);
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__java_optional_filter")),
            args: call_args,
            optional: false,
        });
    }

    if method == "map" && args.len() == 1 && java_optional_receiver(&receiver) {
        let mut call_args = vec![Argument::positional(receiver)];
        call_args.extend(args);
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__java_optional_map")),
            args: call_args,
            optional: false,
        });
    }

    if method == "flatMap" && args.len() == 1 && java_optional_receiver(&receiver) {
        let mut call_args = vec![Argument::positional(receiver)];
        call_args.extend(args);
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__java_optional_flat_map")),
            args: call_args,
            optional: false,
        });
    }

    if method == "ifPresentOrElse" && args.len() == 2 && java_optional_receiver(&receiver) {
        let mut call_args = vec![Argument::positional(receiver)];
        call_args.extend(args);
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__java_optional_if_present_or_else")),
            args: call_args,
            optional: false,
        });
    }

    if method == "isEmpty" && args.is_empty() && java_optional_receiver(&receiver) {
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__java_optional_is_empty")),
            args: vec![Argument::positional(receiver)],
            optional: false,
        });
    }

    if method == "or" && args.len() == 1 && java_optional_receiver(&receiver) {
        let call_supplier = matches!(&args[0].value.kind, ExprKind::Lambda { .. });
        let mut call_args = vec![Argument::positional(receiver)];
        call_args.extend(args);
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident(if call_supplier {
                "__java_optional_or_get"
            } else {
                "__java_optional_or"
            })),
            args: call_args,
            optional: false,
        });
    }

    if method == "stream" && args.is_empty() && java_optional_receiver(&receiver) {
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__java_optional_stream")),
            args: vec![Argument::positional(receiver)],
            optional: false,
        });
    }

    if method == "equals" && args.len() == 1 && java_optional_receiver(&receiver) {
        let mut call_args = vec![Argument::positional(receiver)];
        call_args.extend(args);
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__java_optional_equals")),
            args: call_args,
            optional: false,
        });
    }

    if method == "toString" && args.is_empty() && java_optional_receiver(&receiver) {
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__java_optional_to_string")),
            args: vec![Argument::positional(receiver)],
            optional: false,
        });
    }

    if method == "get" && args.is_empty() && java_optional_receiver(&receiver) {
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__java_stream_optional_get")),
            args: vec![Argument::positional(receiver)],
            optional: false,
        });
    }

    if java_functional_receiver(&receiver) {
        if let Some(expr) = java_functional_default_method(&receiver, method.as_str(), &args) {
            return expr;
        }
    }

    let functional_call_result_receiver = matches!(receiver.kind, ExprKind::Call { .. })
        && java_functional_result_method(method.as_str());
    if (java_functional_receiver(&receiver) || functional_call_result_receiver)
        && java_functional_receiver_method(&receiver, method.as_str())
    {
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Sequence(vec![receiver]))),
            args,
            optional: false,
        });
    }

    if java_list_result_receiver(&receiver) {
        if let Some(internal) = java_list_method_name(method.as_str(), args.len()) {
            let mut call_args = vec![Argument::positional(receiver)];
            call_args.extend(args);
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(internal)),
                args: call_args,
                optional: false,
            });
        }
    }

    if java_map_result_receiver(&receiver) {
        if let Some(internal) = java_map_method_name(method.as_str()) {
            let mut call_args = vec![Argument::positional(receiver)];
            call_args.extend(args);
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(internal)),
                args: call_args,
                optional: false,
            });
        }
    }

    if java_spliterator_receiver(&receiver) {
        if let Some(internal) = java_spliterator_method_name(method.as_str(), args.len()) {
            let mut call_args = vec![Argument::positional(receiver)];
            call_args.extend(args);
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(internal)),
                args: call_args,
                optional: false,
            });
        }
    }

    if java_runtime_receiver(&receiver) {
        if let Some(internal) = java_runtime_method_name(method.as_str(), args.len()) {
            let mut call_args = vec![Argument::positional(receiver)];
            call_args.extend(args);
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(internal)),
                args: call_args,
                optional: false,
            });
        }
    }

    if java_process_builder_receiver(&receiver) {
        if method == "command" && !args.is_empty() {
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__j_pb_command_set")),
                args: vec![
                    Argument::positional(receiver),
                    Argument::positional(java_args_to_array(&args)),
                ],
                optional: false,
            });
        }
        if let Some(internal) = java_process_builder_method_name(method.as_str(), args.len()) {
            let mut call_args = vec![Argument::positional(receiver)];
            call_args.extend(args);
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(internal)),
                args: call_args,
                optional: false,
            });
        }
    }

    if java_process_receiver(&receiver) {
        if let Some(internal) = java_process_method_name(method.as_str(), args.len()) {
            let mut call_args = vec![Argument::positional(receiver)];
            call_args.extend(args);
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(internal)),
                args: call_args,
                optional: false,
            });
        }
    }

    if java_file_receiver(&receiver) && method == "getPath" && args.is_empty() {
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__j_file_get_path")),
            args: vec![Argument::positional(receiver)],
            optional: false,
        });
    }

    if java_redirect_receiver(&receiver) && method == "type" && args.is_empty() {
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__j_pb_redirect_type")),
            args: vec![Argument::positional(receiver)],
            optional: false,
        });
    }

    if matches!(method.as_str(), "wait" | "notify" | "notifyAll") {
        let prelude_fn = match method.as_str() {
            "wait" => "__j_object_wait",
            "notify" => "__j_object_notify",
            _ => "__j_object_notify_all",
        };
        let mut call_args = vec![Argument::positional(receiver)];
        call_args.extend(args);
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident(prelude_fn)),
            args: call_args,
            optional: false,
        });
    }

    if let ExprKind::Ident(name) = &receiver.kind {
        if let Some((class_name, type_name)) = java_current_static_field_type(name) {
            if java_type_is_list_like(Some(&type_name)) {
                if let Some(internal) = java_list_method_name(&method, args.len()) {
                    let mut call_args =
                        vec![Argument::positional(Expression::new(ExprKind::Member {
                            object: Box::new(Expression::ident(&class_name)),
                            field: name.clone(),
                            null_safe: false,
                        }))];
                    call_args.extend(args);
                    return Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident(internal)),
                        args: call_args,
                        optional: false,
                    });
                }
            }
            if java_type_is_queue_or_deque(Some(&type_name)) {
                if let Some(internal) =
                    java_queue_method_name(Some(&type_name), &method, args.len())
                {
                    let mut call_args =
                        vec![Argument::positional(Expression::new(ExprKind::Member {
                            object: Box::new(Expression::ident(&class_name)),
                            field: name.clone(),
                            null_safe: false,
                        }))];
                    call_args.extend(args);
                    return Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident(internal)),
                        args: call_args,
                        optional: false,
                    });
                }
            }
            if java_type_is_semaphore(Some(&type_name)) {
                if let Some(internal) = java_semaphore_method_name(&method) {
                    let mut call_args =
                        vec![Argument::positional(Expression::new(ExprKind::Member {
                            object: Box::new(Expression::ident(&class_name)),
                            field: name.clone(),
                            null_safe: false,
                        }))];
                    call_args.extend(args);
                    return Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident(internal)),
                        args: call_args,
                        optional: false,
                    });
                }
            }
            if java_type_is_count_down_latch(Some(&type_name)) {
                if let Some(internal) = java_count_down_latch_method_name(&method, args.len()) {
                    let mut call_args =
                        vec![Argument::positional(Expression::new(ExprKind::Member {
                            object: Box::new(Expression::ident(&class_name)),
                            field: name.clone(),
                            null_safe: false,
                        }))];
                    call_args.extend(args);
                    return Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident(internal)),
                        args: call_args,
                        optional: false,
                    });
                }
            }
            if java_type_is_future_task(Some(&type_name)) {
                if let Some(internal) = java_future_task_method_name(&method, args.len()) {
                    let mut call_args =
                        vec![Argument::positional(Expression::new(ExprKind::Member {
                            object: Box::new(Expression::ident(&class_name)),
                            field: name.clone(),
                            null_safe: false,
                        }))];
                    call_args.extend(args);
                    return Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident(internal)),
                        args: call_args,
                        optional: false,
                    });
                }
            }
            if java_type_is_executor_service(Some(&type_name)) {
                if let Some(internal) = java_executor_method_name(&method, args.len()) {
                    let mut call_args =
                        vec![Argument::positional(Expression::new(ExprKind::Member {
                            object: Box::new(Expression::ident(&class_name)),
                            field: name.clone(),
                            null_safe: false,
                        }))];
                    call_args.extend(args);
                    return Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident(internal)),
                        args: call_args,
                        optional: false,
                    });
                }
            }
        }
    }

    if java_type_is_list_like(java_static_field_type(&receiver).as_deref()) {
        if let Some(internal) = java_list_method_name(&method, args.len()) {
            let mut call_args = vec![Argument::positional(receiver)];
            call_args.extend(args);
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(internal)),
                args: call_args,
                optional: false,
            });
        }
    }

    if let Some(type_name) = java_static_field_type(&receiver) {
        if java_type_is_queue_or_deque(Some(&type_name)) {
            if let Some(internal) = java_queue_method_name(Some(&type_name), &method, args.len()) {
                let mut call_args = vec![Argument::positional(receiver)];
                call_args.extend(args);
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident(internal)),
                    args: call_args,
                    optional: false,
                });
            }
        }
        if java_type_is_count_down_latch(Some(&type_name)) {
            if let Some(internal) = java_count_down_latch_method_name(&method, args.len()) {
                let mut call_args = vec![Argument::positional(receiver)];
                call_args.extend(args);
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident(internal)),
                    args: call_args,
                    optional: false,
                });
            }
        }
        if java_type_is_future_task(Some(&type_name)) {
            if let Some(internal) = java_future_task_method_name(&method, args.len()) {
                let mut call_args = vec![Argument::positional(receiver)];
                call_args.extend(args);
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident(internal)),
                    args: call_args,
                    optional: false,
                });
            }
        }
        if java_type_is_executor_service(Some(&type_name)) {
            if let Some(internal) = java_executor_method_name(&method, args.len()) {
                let mut call_args = vec![Argument::positional(receiver)];
                call_args.extend(args);
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident(internal)),
                    args: call_args,
                    optional: false,
                });
            }
        }
    }

    let thread_receiver = match &receiver.kind {
        ExprKind::Ident(n) => JAVA_THREAD_VARS.with(|vars| vars.borrow().contains(n.as_str())),
        ExprKind::Call { callee, .. } => {
            matches!(&callee.kind, ExprKind::Ident(n) if matches!(n.as_str(), "__j_thread_current" | "__j_thread_new"))
        }
        _ => false,
    };
    if thread_receiver {
        if method == "start" {
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__j_thread_start")),
                args: vec![Argument::positional(receiver.clone())],
                optional: false,
            });
        }
        let prelude_fn = match method.as_str() {
            "join" => {
                let unsafe_target = match &receiver.kind {
                    ExprKind::Ident(name) => {
                        JAVA_THREAD_UNSAFE_TARGETS.with(|targets| targets.borrow().contains(name))
                    }
                    _ => false,
                };
                Some(if unsafe_target {
                    "__j_thread_join"
                } else {
                    "__java_thread_join"
                })
            }
            "isAlive" => Some("__j_thread_is_alive"),
            "getName" => Some("__j_thread_get_name"),
            "setName" => Some("__j_thread_set_name"),
            "getPriority" => Some("__j_thread_get_priority"),
            "setPriority" => Some("__j_thread_set_priority"),
            "interrupt" => Some("__j_thread_interrupt"),
            "isInterrupted" => Some("__j_thread_is_interrupted"),
            _ => None,
        };
        if let Some(prelude_fn) = prelude_fn {
            let mut call_args = vec![Argument::positional(receiver)];
            call_args.extend(args);
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(prelude_fn)),
                args: call_args,
                optional: false,
            });
        }
    }

    let tlr_receiver = match &receiver.kind {
        ExprKind::Ident(n) => JAVA_TLR_VARS.with(|vars| vars.borrow().contains(n.as_str())),
        ExprKind::Call { callee, .. } => {
            matches!(&callee.kind, ExprKind::Ident(n) if n == "__java_random_new")
        }
        _ => false,
    };
    if tlr_receiver || java_random_receiver(&receiver) {
        let prelude_fn = match method.as_str() {
            "setSeed" => Some("__java_random_set_seed"),
            "nextInt" => Some("__java_random_next_int"),
            "nextLong" => Some("__java_random_next_long"),
            "nextDouble" => Some("__java_random_next_double"),
            "nextFloat" => Some("__java_random_next_float"),
            "nextBoolean" => Some("__java_random_next_boolean"),
            "nextGaussian" => Some("__java_random_next_double"),
            "split" => Some("__java_random_split"),
            "ints" => Some("__java_random_ints"),
            "longs" => Some("__java_random_longs"),
            "doubles" => Some("__java_random_doubles"),
            "nextBytes" => Some("__java_random_next_bytes"),
            _ => None,
        };
        if let Some(prelude_fn) = prelude_fn {
            let mut call_args = vec![Argument::positional(receiver)];
            call_args.extend(args);
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(prelude_fn)),
                args: call_args,
                optional: false,
            });
        }
    }

    // Integer bit ops: wrap the value operand in __j_i32 — the dynamic
    // as_i32 coercion saturates (0x80000000-class literals arrive as
    // f64 2147483648 and clamp to i32::MAX without it).
    if matches!(&receiver.kind, ExprKind::Ident(n) if n == "Integer")
        && matches!(
            method.as_str(),
            "bitCount"
                | "numberOfLeadingZeros"
                | "numberOfTrailingZeros"
                | "rotateLeft"
                | "rotateRight"
                | "lowestOneBit"
                | "highestOneBit"
        )
        && !args.is_empty()
    {
        let mut args = args;
        let value = args.remove(0);
        args.insert(
            0,
            Argument::positional(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__j_i32")),
                args: vec![value],
                optional: false,
            })),
        );
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident("Integer")),
                field: method,
                null_safe: false,
            })),
            args,
            optional: false,
        });
    }

    // Integer.to{Binary,Hex,Octal}String → __j_to_radix (unsigned 32-bit
    // digits; the old common:java.to_*_string arms called ecma:number
    // toBinary/toHex — host fns that never existed).
    if matches!(&receiver.kind, ExprKind::Ident(n) if n == "Integer") && args.len() == 1 {
        let radix = match method.as_str() {
            "toBinaryString" => Some(2),
            "toOctalString" => Some(8),
            "toHexString" => Some(16),
            _ => None,
        };
        if let Some(radix) = radix {
            let mut call_args = args;
            call_args.push(Argument::positional(Expression::int(radix)));
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__j_to_radix")),
                args: call_args,
                optional: false,
            });
        }
    }

    // String.format → the Java Formatter runtime (__j_sprintf), which
    // implements the Java-specific conversions (%b, %,d, %e/%E two-digit
    // exponents, %g, %n) and delegates the rest to the shared engine.
    if method == "format"
        && matches!(&receiver.kind, ExprKind::Ident(n) if n == "String")
        && !args.is_empty()
    {
        let mut it = args.into_iter();
        let fmt = it.next().expect("fmt");
        let rest: Vec<ArrayElement> = it
            .map(|arg| ArrayElement {
                key: None,
                value: arg.value,
                spread: false,
                by_ref: false,
            })
            .collect();
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__j_sprintf")),
            args: vec![
                fmt,
                Argument::positional(Expression::new(ExprKind::Array(rest))),
            ],
            optional: false,
        });
    }
    if let ExprKind::Member {
        object: ref root_obj,
        field: ref root_field,
        ..
    } = receiver.kind
    {
        if let ExprKind::Ident(ref root_name) = root_obj.kind {
            // System.exit(code) → __process_exit(code)
            if root_name == "System" && root_field == "exit" {
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__process_exit")),
                    args,
                    optional: false,
                });
            }
        }
    }
    if method == "of"
        && java_expr_dotted_name(&receiver).as_deref() == Some("Character.UnicodeBlock")
    {
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__j_char_unicode_block_of")),
            args,
            optional: false,
        });
    }
    // StringBuilder/StringBuffer receivers → the __j_sb_* runtime
    // (emitter/format_runtime.rs; the builder stores its text in
    // `__buffer`). Read-only string queries re-receiver onto the buffer
    // string so the ordinary string value-methods dispatch fires;
    // mutators call the prelude fns (which return the builder — JLS
    // `this` — so chains stay sb-shaped).
    let sb_receiver = match &receiver.kind {
        ExprKind::Ident(n) => JAVA_SB_VARS.with(|vars| vars.borrow().contains(n.as_str())),
        ExprKind::New { class, .. } => matches!(
            &class.kind,
            ExprKind::Ident(c) if c == "StringBuilder" || c == "StringBuffer"
        ),
        ExprKind::Call { callee, .. } => matches!(
            &callee.kind,
            ExprKind::Ident(n) if matches!(
                n.as_str(),
                "__j_sb_append"
                    | "__j_sb_insert"
                    | "__j_sb_delete"
                    | "__j_sb_delete_char_at"
                    | "__j_sb_replace"
                    | "__j_sb_reverse"
            )
        ),
        _ => false,
    };
    if sb_receiver {
        let prelude_fn = match method.as_str() {
            "toString" => Some("__j_sb_to_string"),
            "length" => Some("__j_sb_length"),
            "append" => Some("__j_sb_append"),
            "charAt" => Some("__j_sb_char_at"),
            "setCharAt" => Some("__j_sb_set_char_at"),
            "insert" => Some("__j_sb_insert"),
            "delete" => Some("__j_sb_delete"),
            "deleteCharAt" => Some("__j_sb_delete_char_at"),
            "replace" if args.len() == 3 => Some("__j_sb_replace"),
            "reverse" => Some("__j_sb_reverse"),
            "setLength" => Some("__j_sb_set_length"),
            "capacity" => Some("__j_sb_capacity"),
            "ensureCapacity" => Some("__j_sb_ensure_capacity"),
            "compareTo" => Some("__j_sb_compare_to"),
            "appendCodePoint" => Some("__j_sb_append_code_point"),
            _ => None,
        };
        if let Some(prelude_fn) = prelude_fn {
            let mut call_args = vec![Argument::positional(receiver)];
            call_args.extend(args);
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(prelude_fn)),
                args: call_args,
                optional: false,
            });
        }
        if matches!(
            method.as_str(),
            "substring"
                | "indexOf"
                | "lastIndexOf"
                | "codePointAt"
                | "codePointCount"
                | "isEmpty"
                | "contains"
        ) {
            let as_string = Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__j_sb_to_string")),
                args: vec![Argument::positional(receiver)],
                optional: false,
            });
            if method == "codePointCount" {
                let mut call_args = vec![Argument::positional(as_string)];
                call_args.extend(args);
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__j_string_code_point_count")),
                    args: call_args,
                    optional: false,
                });
            }
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(as_string),
                    field: method,
                    null_safe: false,
                })),
                args,
                optional: false,
            });
        }
    }

    let bigint_receiver = match &receiver.kind {
        ExprKind::Ident(n) => JAVA_BIGINT_VARS.with(|vars| vars.borrow().contains(n.as_str())),
        ExprKind::New { class, .. } => match &class.kind {
            ExprKind::Ident(c) => java_type_simple_name(c) == "BigInteger",
            ExprKind::Member { .. } => java_qualified_static_type(class)
                .is_some_and(|name| java_type_simple_name(&name) == "BigInteger"),
            _ => false,
        },
        ExprKind::Call { callee, .. } => matches!(
            &callee.kind,
            ExprKind::Ident(n) if n.starts_with("__java_bigint") || n == "__java_bigint"
        ),
        _ => java_bigint_constant_replacement(&receiver).is_some(),
    };
    if bigint_receiver {
        if let Some(prelude_fn) = java_bigint_method_name(&method) {
            let mut call_args = vec![Argument::positional(receiver)];
            call_args.extend(args);
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(prelude_fn)),
                args: call_args,
                optional: false,
            });
        }
    }

    let bigdecimal_receiver = match &receiver.kind {
        ExprKind::Ident(n) => JAVA_BIGDECIMAL_VARS.with(|vars| vars.borrow().contains(n.as_str())),
        ExprKind::New { class, .. } => match &class.kind {
            ExprKind::Ident(c) => java_type_simple_name(c) == "BigDecimal",
            ExprKind::Member { .. } => java_qualified_static_type(class)
                .is_some_and(|name| java_type_simple_name(&name) == "BigDecimal"),
            _ => false,
        },
        ExprKind::Call { callee, .. } => matches!(
            &callee.kind,
            ExprKind::Ident(n) if java_bigdecimal_function_returns_bigdecimal(n)
        ),
        _ => java_bigdecimal_constant_replacement(&receiver).is_some(),
    };
    if bigdecimal_receiver {
        if let Some(prelude_fn) = java_bigdecimal_method_name(&method, args.len()) {
            let mut call_args = vec![Argument::positional(receiver)];
            call_args.extend(args);
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(prelude_fn)),
                args: call_args,
                optional: false,
            });
        }
    }

    let decimal_format_receiver = match &receiver.kind {
        ExprKind::Ident(n) => {
            JAVA_DECIMAL_FORMAT_VARS.with(|vars| vars.borrow().contains(n.as_str()))
        }
        ExprKind::Call { callee, .. } => matches!(
            &callee.kind,
            ExprKind::Ident(n) if matches!(n.as_str(), "__j_df_new" | "__j_df_currency" | "__j_df_percent" | "__j_df_clone")
        ),
        _ => false,
    };
    if decimal_format_receiver {
        if let Some(prelude_fn) = java_decimal_format_method_name(&method) {
            let mut call_args = vec![Argument::positional(receiver)];
            call_args.extend(args);
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(prelude_fn)),
                args: call_args,
                optional: false,
            });
        }
    }

    if method == "doubleValue"
        && args.is_empty()
        && matches!(&receiver.kind, ExprKind::Ident(n) if JAVA_NUMBER_VARS.with(|vars| vars.borrow().contains(n.as_str())))
    {
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__java_double_to_string")),
            args: vec![Argument::positional(receiver)],
            optional: false,
        });
    }

    if java_stream_builder_receiver(&receiver) {
        if method == "build" && args.is_empty() {
            return receiver;
        }
        if method == "add" && args.len() == 1 {
            let mut call_args = vec![Argument::positional(receiver)];
            call_args.extend(args);
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__j_stream_builder_add")),
                args: call_args,
                optional: false,
            });
        }
    }

    if method == "isParallel" && args.is_empty() {
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__j_stream_is_parallel")),
            args: vec![Argument::positional(receiver)],
            optional: false,
        });
    }

    if method == "compareTo"
        && (matches!(&receiver.kind, ExprKind::Lit(Literal::Str(_)))
            || matches!(&receiver.kind, ExprKind::Ident(n) if JAVA_STRING_VARS.with(|vars| vars.borrow().contains(n.as_str())))
            || matches!(&receiver.kind, ExprKind::New { class, .. } if matches!(&class.kind, ExprKind::Ident(name) if name == "String"))
            || matches!(&receiver.kind, ExprKind::Call { callee, .. } if matches!(&callee.kind, ExprKind::Ident(name) if name == "__java_string_value_of" || name == "__j_sb_to_string")))
    {
        let mut call_args = vec![Argument::positional(receiver)];
        call_args.extend(args);
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__j_string_compare_to")),
            args: call_args,
            optional: false,
        });
    }

    let string_prelude_fn = match method.as_str() {
        "split" if args.len() == 2 => Some("__j_string_split_n"),
        "split" => Some("__j_string_split"),
        "lines" if args.is_empty() => Some("__j_str_lines"),
        "codePointBefore" => Some("__j_string_code_point_before"),
        "codePointCount" => Some("__j_string_code_point_count"),
        "offsetByCodePoints" => Some("__j_string_offset_by_code_points"),
        "regionMatches" if args.len() == 4 => Some("__j_string_region_matches"),
        "regionMatches" if args.len() == 5 => Some("__j_string_region_matches_ignore"),
        "getBytes" => Some("__j_string_get_bytes"),
        "chars" => Some("__j_string_chars"),
        "codePoints" => Some("__j_string_code_points"),
        "stripIndent" => Some("__j_string_strip_indent"),
        "translateEscapes" => Some("__j_string_translate_escapes"),
        _ => None,
    };
    if method == "length" && args.is_empty() {
        if let ExprKind::Call {
            callee,
            args: call_args,
            ..
        } = &receiver.kind
        {
            if matches!(&callee.kind, ExprKind::Ident(name) if name == "__j_string_from_array")
                && call_args.len() == 1
            {
                return Expression::new(ExprKind::Member {
                    object: Box::new(call_args[0].value.clone()),
                    field: "length".to_string(),
                    null_safe: false,
                });
            }
        }
        if let ExprKind::New {
            class,
            args: ctor_args,
        } = &receiver.kind
        {
            if matches!(&class.kind, ExprKind::Ident(name) if name == "String")
                && ctor_args.len() == 1
            {
                let source = &ctor_args[0];
                let is_char_source = java_arg_is_char_array(source)
                    || matches!(
                        &source.value.kind,
                        ExprKind::Call { callee, .. }
                            if matches!(&callee.kind, ExprKind::Ident(name) if name == "__j_char_to_chars")
                    );
                if is_char_source {
                    return Expression::new(ExprKind::Member {
                        object: Box::new(source.value.clone()),
                        field: "length".to_string(),
                        null_safe: false,
                    });
                }
            }
        }
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__j_str_length")),
            args: vec![Argument::positional(receiver)],
            optional: false,
        });
    }
    let receiver_is_static_type = match &receiver.kind {
        ExprKind::Ident(name) => is_java_type_or_util(name),
        _ => java_qualified_static_type(&receiver).is_some(),
    };
    // Pattern.split is java.util.regex semantics, not String.split — leave
    // it for the Pattern-receiver arm below.
    let receiver_is_pattern = match &receiver.kind {
        ExprKind::Ident(n) => JAVA_PATTERN_VARS.with(|vars| vars.borrow().contains(n.as_str())),
        ExprKind::Call { callee, .. } => {
            matches!(&callee.kind, ExprKind::Ident(n) if n == "__j_pat_compile")
        }
        _ => false,
    };
    if !receiver_is_static_type && !receiver_is_pattern {
        if let Some(prelude_fn) = string_prelude_fn {
            let mut call_args = vec![Argument::positional(receiver)];
            call_args.extend(args);
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(prelude_fn)),
                args: call_args,
                optional: false,
            });
        }
    }

    let string_joiner_receiver = match &receiver.kind {
        ExprKind::Ident(n) => {
            JAVA_STRING_JOINER_VARS.with(|vars| vars.borrow().contains(n.as_str()))
        }
        ExprKind::Call { callee, .. } => {
            matches!(&callee.kind, ExprKind::Ident(n) if matches!(n.as_str(), "__j_sj_new" | "__j_sj_add" | "__j_sj_merge" | "__j_sj_set_empty_value"))
        }
        _ => false,
    };
    if string_joiner_receiver {
        let prelude_fn = match method.as_str() {
            "add" => Some("__j_sj_add"),
            "merge" => Some("__j_sj_merge"),
            "setEmptyValue" => Some("__j_sj_set_empty_value"),
            "length" => Some("__j_sj_length"),
            "toString" => Some("__j_sj_to_string"),
            _ => None,
        };
        if let Some(prelude_fn) = prelude_fn {
            let mut call_args = vec![Argument::positional(receiver)];
            call_args.extend(args);
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(prelude_fn)),
                args: call_args,
                optional: false,
            });
        }
    }

    let string_tokenizer_receiver = match &receiver.kind {
        ExprKind::Ident(n) => {
            JAVA_STRING_TOKENIZER_VARS.with(|vars| vars.borrow().contains(n.as_str()))
        }
        ExprKind::Call { callee, .. } => {
            matches!(&callee.kind, ExprKind::Ident(n) if n == "__j_st_new")
        }
        _ => false,
    };
    if string_tokenizer_receiver {
        let prelude_fn = match method.as_str() {
            "hasMoreTokens" | "hasMoreElements" => Some("__j_st_has_more"),
            "nextToken" | "nextElement" => Some("__j_st_next"),
            "countTokens" => Some("__j_st_count"),
            _ => None,
        };
        if let Some(prelude_fn) = prelude_fn {
            let mut call_args = vec![Argument::positional(receiver)];
            call_args.extend(args);
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(prelude_fn)),
                args: call_args,
                optional: false,
            });
        }
    }

    let scanner_receiver = match &receiver.kind {
        ExprKind::Ident(n) => JAVA_SCANNER_VARS.with(|vars| vars.borrow().contains(n.as_str())),
        ExprKind::Call { callee, .. } => {
            matches!(&callee.kind, ExprKind::Ident(n) if matches!(n.as_str(), "__j_sc_new" | "__j_sc_use_delim" | "__j_sc_use_locale" | "__j_sc_use_radix"))
        }
        _ => false,
    };
    if scanner_receiver {
        let prelude_fn = match method.as_str() {
            "next" => Some("__j_sc_next"),
            "nextInt" => Some("__j_sc_next_int"),
            "nextLong" => Some("__j_sc_next_long"),
            "nextDouble" | "nextFloat" | "nextBigDecimal" => Some("__j_sc_next_double"),
            "nextBoolean" => Some("__j_sc_next_bool"),
            "nextLine" => Some("__j_sc_next_line"),
            "hasNext" => Some("__j_sc_has_next"),
            "hasNextInt" | "hasNextLong" => Some("__j_sc_has_next_int"),
            "hasNextDouble" | "hasNextBigDecimal" => Some("__j_sc_has_next_double"),
            "hasNextLine" => Some("__j_sc_has_next_line"),
            "useDelimiter" => Some("__j_sc_use_delim"),
            "useLocale" => Some("__j_sc_use_locale"),
            "useRadix" => Some("__j_sc_use_radix"),
            "skip" => Some("__j_sc_skip"),
            "findInLine" | "findWithinHorizon" => Some("__j_sc_find"),
            "close" => Some("__j_sc_close"),
            _ => None,
        };
        if let Some(prelude_fn) = prelude_fn {
            let mut call_args = vec![Argument::positional(receiver)];
            call_args.extend(args);
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(prelude_fn)),
                args: call_args,
                optional: false,
            });
        }
    }

    let formatter_receiver = match &receiver.kind {
        ExprKind::Ident(n) => JAVA_FORMATTER_VARS.with(|vars| vars.borrow().contains(n.as_str())),
        ExprKind::Call { callee, .. } => {
            matches!(&callee.kind, ExprKind::Ident(n) if matches!(n.as_str(), "__j_fmt_new" | "__j_fmt_format"))
        }
        _ => false,
    };
    if formatter_receiver {
        let prelude_fn = match method.as_str() {
            "format" => Some("__j_fmt_format"),
            "toString" => Some("__j_fmt_to_string"),
            "locale" => Some("__j_fmt_locale"),
            "out" => Some("__j_fmt_out"),
            _ => None,
        };
        if let Some(prelude_fn) = prelude_fn {
            let mut call_args = vec![Argument::positional(receiver)];
            if method == "format" && !args.is_empty() {
                let mut it = args.into_iter();
                let fmt = it.next().expect("fmt");
                let rest: Vec<ArrayElement> = it
                    .map(|arg| ArrayElement {
                        key: None,
                        value: arg.value,
                        spread: false,
                        by_ref: false,
                    })
                    .collect();
                call_args.push(fmt);
                call_args.push(Argument::positional(Expression::new(ExprKind::Array(rest))));
            } else {
                call_args.extend(args);
            }
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(prelude_fn)),
                args: call_args,
                optional: false,
            });
        }
    }

    let message_format_receiver = match &receiver.kind {
        ExprKind::Ident(n) => {
            JAVA_MESSAGE_FORMAT_VARS.with(|vars| vars.borrow().contains(n.as_str()))
        }
        ExprKind::Call { callee, .. } => {
            matches!(&callee.kind, ExprKind::Ident(n) if matches!(n.as_str(), "__j_mf_new" | "__j_mf_clone"))
        }
        _ => false,
    };
    if message_format_receiver {
        let prelude_fn = match method.as_str() {
            "format" => Some("__j_mf_format"),
            "applyPattern" => Some("__j_mf_apply_pattern"),
            "toPattern" => Some("__j_mf_to_pattern"),
            "setLocale" => Some("__j_mf_set_locale"),
            "parse" => Some("__j_mf_parse"),
            "clone" => Some("__j_mf_clone"),
            "equals" => Some("__j_mf_equals"),
            _ => None,
        };
        if let Some(prelude_fn) = prelude_fn {
            let mut call_args = vec![Argument::positional(receiver)];
            call_args.extend(args);
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(prelude_fn)),
                args: call_args,
                optional: false,
            });
        }
    }

    let calendar_receiver = match &receiver.kind {
        ExprKind::Ident(n) => JAVA_CALENDAR_VARS.with(|vars| vars.borrow().contains(n.as_str())),
        ExprKind::Call { callee, .. } => {
            matches!(&callee.kind, ExprKind::Ident(n) if n == "__j_cal_new")
        }
        _ => false,
    };
    if calendar_receiver {
        let prelude_fn = match method.as_str() {
            "set" => Some("__j_cal_set"),
            "getTime" => Some("__j_cal_get_time"),
            _ => None,
        };
        if let Some(prelude_fn) = prelude_fn {
            let mut call_args = vec![Argument::positional(receiver)];
            call_args.extend(args);
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(prelude_fn)),
                args: call_args,
                optional: false,
            });
        }
    }

    // java.util.regex — Pattern.compile plus Pattern/Matcher instance
    // methods route through the __j_pat_*/__j_m_* prelude runtime.
    {
        let recv_dotted = java_expr_dotted_name(&receiver);
        let is_pattern_type = matches!(
            recv_dotted.as_deref(),
            Some("Pattern") | Some("java.util.regex.Pattern")
        );
        if is_pattern_type && method == "compile" && !args.is_empty() {
            if args.len() >= 2 {
                let mut it = args.into_iter();
                let re = it.next().expect("regex");
                let flags = it.next().expect("flags");
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__j_pat_compile_flags")),
                    args: vec![re, flags],
                    optional: false,
                });
            }
            let re = args.into_iter().next().expect("regex");
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__j_pat_compile")),
                args: vec![re],
                optional: false,
            });
        }
        let pattern_recv = match &receiver.kind {
            ExprKind::Ident(n) => JAVA_PATTERN_VARS.with(|vars| vars.borrow().contains(n.as_str())),
            ExprKind::Call { callee, .. } => {
                matches!(&callee.kind, ExprKind::Ident(n) if n == "__j_pat_compile")
            }
            _ => false,
        };
        if pattern_recv {
            let prelude_fn = match method.as_str() {
                "matcher" => Some("__j_pat_matcher"),
                "split" if args.len() == 2 => Some("__j_pat_split_n"),
                "split" => Some("__j_pat_split"),
                "pattern" | "toString" => Some("__j_pat_pattern"),
                "flags" => Some("__j_pat_flags"),
                _ => None,
            };
            if let Some(prelude_fn) = prelude_fn {
                let mut call_args = vec![Argument::positional(receiver)];
                call_args.extend(args);
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident(prelude_fn)),
                    args: call_args,
                    optional: false,
                });
            }
        }
        let matcher_recv = match &receiver.kind {
            ExprKind::Ident(n) => JAVA_MATCHER_VARS.with(|vars| vars.borrow().contains(n.as_str())),
            ExprKind::Call { callee, .. } => {
                matches!(&callee.kind, ExprKind::Ident(n) if n == "__j_pat_matcher")
            }
            _ => false,
        };
        if matcher_recv {
            let prelude_fn = match method.as_str() {
                "find" => Some("__j_m_find"),
                "matches" => Some("__j_m_matches"),
                "lookingAt" => Some("__j_m_looking_at"),
                "group" => Some("__j_m_group"),
                "replaceAll" => Some("__j_m_replace_all"),
                "replaceFirst" => Some("__j_m_replace_first"),
                "appendReplacement" => Some("__j_m_append_replacement"),
                "appendTail" => Some("__j_m_append_tail"),
                "reset" => Some("__j_m_reset"),
                _ => None,
            };
            if let Some(prelude_fn) = prelude_fn {
                let mut call_args = vec![Argument::positional(receiver)];
                if method == "group" && args.is_empty() {
                    // group() is group(0) — the whole match.
                    call_args.push(Argument::positional(Expression::int(0)));
                }
                call_args.extend(args);
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident(prelude_fn)),
                    args: call_args,
                    optional: false,
                });
            }
        }
    }

    // javax.xml.namespace.QName receivers → common XML name accessors.
    {
        let qname_recv = match &receiver.kind {
            ExprKind::Ident(n) => JAVA_QNAME_VARS.with(|vars| vars.borrow().contains(n.as_str())),
            ExprKind::Call { callee, .. } => matches!(
                &callee.kind,
                ExprKind::Ident(n) if matches!(n.as_str(), "__java_xml_name" | "__java_xml_node_name")
            ),
            _ => false,
        };
        if qname_recv {
            let prelude_fn = match method.as_str() {
                "getLocalPart" => Some("__java_xml_local"),
                "getNamespaceURI" => Some("__java_xml_namespace"),
                "getPrefix" => Some("__java_xml_prefix"),
                "toString" => Some("__java_xml_qualified"),
                "equals" if args.len() == 1 => Some("__java_xml_equal"),
                _ => None,
            };
            if let Some(prelude_fn) = prelude_fn {
                let mut call_args = vec![Argument::positional(receiver)];
                call_args.extend(args);
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident(prelude_fn)),
                    args: call_args,
                    optional: false,
                });
            }
        }
    }

    // java.net.URL/URI receivers → the __j_url_* getters over the
    // WHATWG-parsed object. toURI()/toURL() are identity (same object).
    {
        let url_recv = match &receiver.kind {
            ExprKind::Ident(n) => JAVA_URL_VARS.with(|vars| vars.borrow().contains(n.as_str())),
            ExprKind::Call { callee, .. } => matches!(
                &callee.kind,
                ExprKind::Ident(n) if matches!(
                    n.as_str(),
                    "__j_url_new" | "__j_url_ctx" | "__j_url_make" | "__j_url_parse"
                        | "__j_uri_new"
                        | "__j_uri_make3"
                        | "__j_uri_make7"
                        | "__j_uri_normalize"
                        | "__j_uri_resolve"
                        | "__j_uri_relativize"
                )
            ),
            _ => false,
        };
        if url_recv {
            if matches!(method.as_str(), "toURI" | "toURL") {
                return receiver;
            }
            let prelude_fn = match method.as_str() {
                "getProtocol" | "getScheme" => Some("__j_url_protocol"),
                "getHost" => Some("__j_url_host"),
                "getPort" => Some("__j_url_port"),
                "getDefaultPort" => Some("__j_url_default_port"),
                "getPath" | "getRawPath" => Some("__j_url_path"),
                "getQuery" | "getRawQuery" => Some("__j_url_query"),
                "getRef" | "getFragment" | "getRawFragment" => Some("__j_url_ref"),
                "getFile" => Some("__j_url_file"),
                "getAuthority" | "getRawAuthority" => Some("__j_url_authority"),
                "getUserInfo" | "getRawUserInfo" => Some("__j_url_user_info"),
                "getSchemeSpecificPart" | "getRawSchemeSpecificPart" => Some("__j_uri_ssp"),
                "isAbsolute" => Some("__j_uri_is_absolute"),
                "isOpaque" => Some("__j_uri_is_opaque"),
                "normalize" => Some("__j_uri_normalize"),
                "resolve" => Some("__j_uri_resolve"),
                "relativize" => Some("__j_uri_relativize"),
                "compareTo" => Some("__j_uri_compare_to"),
                "toString" | "toExternalForm" | "toASCIIString" => Some("__j_url_to_string"),
                "equals" => Some("__j_url_equals"),
                "hashCode" => Some("__j_url_hash"),
                "sameFile" => Some("__j_url_same_file"),
                _ => None,
            };
            if let Some(prelude_fn) = prelude_fn {
                let mut call_args = vec![Argument::positional(receiver)];
                call_args.extend(args);
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident(prelude_fn)),
                    args: call_args,
                    optional: false,
                });
            }
        }
    }

    // java.util.Base64 encoder/decoder objects — small Java prelude over
    // ECMA btoa/atob.
    {
        let b64_recv = match &receiver.kind {
            ExprKind::Call { callee, .. } => matches!(
                &callee.kind,
                ExprKind::Ident(n) if matches!(
                    n.as_str(),
                    "__j_b64_encoder"
                        | "__j_b64_url_encoder"
                        | "__j_b64_mime_encoder"
                        | "__j_b64_decoder"
                        | "__j_b64_url_decoder"
                        | "__j_b64_mime_decoder"
                        | "__j_b64_without_padding"
                )
            ),
            _ => false,
        };
        if b64_recv {
            let prelude_fn = match method.as_str() {
                "withoutPadding" => Some("__j_b64_without_padding"),
                "encodeToString" => Some("__j_b64_encode_to_string"),
                "encode" => Some("__j_b64_encode"),
                "decode" => Some("__j_b64_decode"),
                _ => None,
            };
            if let Some(prelude_fn) = prelude_fn {
                let mut call_args = vec![Argument::positional(receiver)];
                call_args.extend(args);
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident(prelude_fn)),
                    args: call_args,
                    optional: false,
                });
            }
        }
    }

    {
        let prelude_fn = match &receiver.kind {
            ExprKind::Call { callee, .. }
                if matches!(&callee.kind, ExprKind::Ident(n) if n == "__j_props_enum")
                    && matches!(method.as_str(), "hasMoreElements" | "hasMoreTokens") =>
            {
                Some("__j_enum_has_more")
            }
            ExprKind::Call { callee, .. }
                if matches!(&callee.kind, ExprKind::Ident(n) if n == "__j_props_class")
                    && method == "getName" =>
            {
                Some("__j_class_get_name")
            }
            ExprKind::Call { callee, .. }
                if matches!(&callee.kind, ExprKind::Ident(n) if n == "__j_system_get_properties")
                    && method == "getClass" =>
            {
                Some("__j_props_class")
            }
            _ => None,
        };
        if let Some(prelude_fn) = prelude_fn {
            let mut call_args = vec![Argument::positional(receiver)];
            call_args.extend(args);
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(prelude_fn)),
                args: call_args,
                optional: false,
            });
        }
    }

    // Throwable accessors: the canonical exception object carries the
    // ECMA Error shape — `message` is a property, not a method.
    if matches!(method.as_str(), "getMessage" | "getLocalizedMessage") && args.is_empty() {
        return Expression::new(ExprKind::Member {
            object: Box::new(receiver),
            field: "message".to_string(),
            null_safe: false,
        });
    }
    if method == "getCause" && args.is_empty() {
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__j_get_cause")),
            args: vec![Argument::positional(receiver)],
            optional: false,
        });
    }
    if method == "initCause" && args.len() == 1 {
        let mut call_args = vec![Argument::positional(receiver)];
        call_args.extend(args);
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__j_init_cause")),
            args: call_args,
            optional: false,
        });
    }

    // System.arraycopy(src, srcPos, dest, destPos, len) →
    // __j_arraycopy prelude fn (JLS in-place, overlap-safe).
    if method == "arraycopy" && matches!(&receiver.kind, ExprKind::Ident(n) if n == "System") {
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__j_arraycopy")),
            args,
            optional: false,
        });
    }

    if java_expr_dotted_name(&receiver).as_deref() == Some("java.math.BigInteger")
        && method == "valueOf"
    {
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__java_bigint")),
            args,
            optional: false,
        });
    }

    if let Some(type_name) = java_expr_dotted_name(&receiver) {
        if java_type_simple_name(&type_name) == "URI" {
            if let Some(expr) = java_uri_static_call(method.as_str(), args.clone()) {
                return expr;
            }
        }
        if java_type_simple_name(&type_name) == "StreamSupport" && method == "stream" {
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__j_stream_support_stream")),
                args,
                optional: false,
            });
        }
        if java_type_simple_name(&type_name) == "Runtime" && method == "getRuntime" {
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__j_runtime_get")),
                args,
                optional: false,
            });
        }
    }

    // `List.remove(int)` and `List.remove(Object)` are distinct Java overloads.
    // The boxed form is explicit in the parsed tree, so preserve that distinction
    // before profile dispatch erases the receiver type.
    if method == "remove"
        && args.len() == 1
        && matches!(
            args[0].value.kind,
            ExprKind::Call { ref callee, .. }
                if matches!(callee.kind, ExprKind::Ident(ref name) if name == "Integer.valueOf")
        )
    {
        let mut call_args = Vec::with_capacity(2);
        call_args.push(Argument::positional(receiver));
        call_args.extend(args);
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__java_list_remove_value")),
            args: call_args,
            optional: false,
        });
    }

    if let Some(type_name) = java_qualified_static_type(&receiver) {
        if let Some(expr) = java_functional_static_call(&type_name, method.as_str(), &args) {
            return expr;
        }
        if type_name == "BigInteger" && method == "valueOf" {
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__java_bigint")),
                args,
                optional: false,
            });
        }
        if java_type_simple_name(&type_name) == "BigDecimal" && method == "valueOf" {
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__j_bd_new")),
                args,
                optional: false,
            });
        }
        if java_type_simple_name(&type_name) == "Optional"
            && method == "of"
            && args.len() == 1
            && java_expr_is_long_value(&args[0].value)
        {
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__java_optional_of_long")),
                args,
                optional: false,
            });
        }
        if java_type_simple_name(&type_name) == "NumberFormat" && method == "getCurrencyInstance" {
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__j_df_currency")),
                args,
                optional: false,
            });
        }
        if java_type_simple_name(&type_name) == "NumberFormat" && method == "getPercentInstance" {
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__j_df_percent")),
                args,
                optional: false,
            });
        }
        if java_type_simple_name(&type_name) == "MessageFormat" && method == "format" {
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__j_mf_static_format")),
                args,
                optional: false,
            });
        }
        if java_type_simple_name(&type_name) == "TimeZone" && method == "getTimeZone" {
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__j_tz_get")),
                args,
                optional: false,
            });
        }
        if java_type_simple_name(&type_name) == "Objects" {
            if let Some(expr) = java_objects_static_call(method.as_str(), args.clone()) {
                return expr;
            }
        }
        if type_name == "StrictMath" && method == "copySign" {
            if java_args_are_copy_sign_negative_zero(&args) {
                return Expression::string("-0.0");
            }
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__java_double_to_string")),
                args: vec![Argument::positional(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("StrictMath.copySign")),
                    args,
                    optional: false,
                }))],
                optional: false,
            });
        }
        if type_name == "Comparator" {
            if let Some(expr) = normalise_comparator_static_call(&method, args.clone()) {
                return expr;
            }
        }
        if type_name == "Double" {
            if let Some(callee) = java_double_static_prelude_fn(&method) {
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident(callee)),
                    args,
                    optional: false,
                });
            }
        }
        if type_name == "Base64" {
            if let Some(expr) = java_base64_static_call(&method, args.clone()) {
                return expr;
            }
        }
        if type_name == "System" {
            if let Some(expr) = java_system_static_call(&method, args.clone()) {
                return expr;
            }
        }
        if type_name == "Thread" {
            if let Some(expr) = java_thread_static_call(&method, args.clone()) {
                return expr;
            }
        }
        if type_name == "ThreadLocalRandom" {
            if method == "current" && args.is_empty() {
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__java_random_new")),
                    args,
                    optional: false,
                });
            }
        }
        if java_type_simple_name(&type_name) == "StreamSupport" && method == "stream" {
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__j_stream_support_stream")),
                args,
                optional: false,
            });
        }
        if java_type_simple_name(&type_name) == "Runtime" && method == "getRuntime" {
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__j_runtime_get")),
                args,
                optional: false,
            });
        }
        if java_type_simple_name(&type_name) == "URI" {
            if let Some(expr) = java_uri_static_call(method.as_str(), args.clone()) {
                return expr;
            }
        }
        if type_name == "Boolean" && matches!(method.as_str(), "parseBoolean" | "valueOf") {
            return java_boolean_parse_call(args);
        }
        if type_name == "Integer" && method == "toString" && !args.is_empty() {
            return java_integer_to_string_call(args);
        }
        if type_name == "String" && method == "valueOf" {
            let callee = if args.len() == 3 {
                "__j_array_chars_to_string"
            } else if args.len() == 1 && java_arg_is_char_array(&args[0]) {
                "__j_string_copy_value_of"
            } else {
                "__java_string_value_of"
            };
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(callee)),
                args,
                optional: false,
            });
        }
        if type_name == "String" && method == "format" {
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__java_string_format")),
                args,
                optional: false,
            });
        }
        if type_name == "String" && method == "join" {
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__java_string_join")),
                args,
                optional: false,
            });
        }
        if type_name == "String" && matches!(method.as_str(), "copyValueOf" | "translateEscapes") {
            let callee = if method == "copyValueOf" && args.len() == 3 {
                "__j_array_chars_to_string"
            } else if method == "copyValueOf" {
                "__j_string_copy_value_of"
            } else {
                "__j_string_translate_escapes"
            };
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(callee)),
                args,
                optional: false,
            });
        }
        if type_name == "Character" {
            if let Some(callee) = java_character_prelude_fn(&method) {
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident(callee)),
                    args,
                    optional: false,
                });
            }
        }
        if type_name == "Character.UnicodeBlock" && method == "of" {
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__j_char_unicode_block_of")),
                args,
                optional: false,
            });
        }
        if type_name.contains('.') {
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(Expression::ident(&type_name)),
                    field: method,
                    null_safe: false,
                })),
                args,
                optional: false,
            });
        }
        let dotted = format!("{}.{}", type_name, method);
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident(&dotted)),
            args,
            optional: false,
        });
    }

    // Static type method calls: Integer.parseInt("42") → call "Integer.parseInt"
    // The profile has dotted builtins like "Integer.parseInt", "Math.max", etc.
    if let ExprKind::Ident(ref type_name) = receiver.kind {
        if is_java_type_or_util(type_name) {
            // `Class.forName(name)` is a runtime type lookup, not a string.
            // A JDK type keeps the existing pass-through (there is no compiled
            // global for `java.lang.String`, and its reflection surface is
            // served off the qualified name). Anything else — a user class, or
            // a name that resolves to nothing — goes through the shared
            // dynamic-symbol path, which consults any registered resolver and
            // then raises `ClassNotFoundException`. Knowing which names are
            // JDK types is Java's business and stays here.
            if type_name == "Class" && method == "forName" && args.len() == 1 {
                let jdk_type = match &args[0].value.kind {
                    ExprKind::Lit(Literal::Str(name)) => {
                        is_java_type_or_util(name.rsplit('.').next().unwrap_or(name))
                    }
                    _ => false,
                };
                if jdk_type {
                    return args[0].value.clone();
                }
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__vybe_class_for_name")),
                    args: vec![Argument {
                        name: None,
                        value: args[0].value.clone(),
                        by_ref: false,
                        spread: false,
                    }],
                    optional: false,
                });
            }
            if let Some(expr) = java_functional_static_call(type_name, method.as_str(), &args) {
                return expr;
            }
            if type_name == "Comparator" {
                if let Some(expr) = normalise_comparator_static_call(&method, args.clone()) {
                    return expr;
                }
            }
            if type_name == "Double" {
                if let Some(callee) = java_double_static_prelude_fn(&method) {
                    return Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident(callee)),
                        args,
                        optional: false,
                    });
                }
            }
            if type_name == "Base64" {
                if let Some(expr) = java_base64_static_call(&method, args.clone()) {
                    return expr;
                }
            }
            if type_name == "System" {
                if let Some(expr) = java_system_static_call(&method, args.clone()) {
                    return expr;
                }
            }
            if type_name == "Thread" {
                if let Some(expr) = java_thread_static_call(&method, args.clone()) {
                    return expr;
                }
            }
            if type_name == "ThreadLocalRandom" {
                if method == "current" && args.is_empty() {
                    return Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__java_random_new")),
                        args,
                        optional: false,
                    });
                }
            }
            if java_type_simple_name(type_name) == "StreamSupport" && method == "stream" {
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__j_stream_support_stream")),
                    args,
                    optional: false,
                });
            }
            if java_type_simple_name(type_name) == "Runtime" && method == "getRuntime" {
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__j_runtime_get")),
                    args,
                    optional: false,
                });
            }
            if java_type_simple_name(type_name) == "URI" {
                if let Some(expr) = java_uri_static_call(method.as_str(), args.clone()) {
                    return expr;
                }
            }
            if type_name == "Boolean" && matches!(method.as_str(), "parseBoolean" | "valueOf") {
                return java_boolean_parse_call(args);
            }
            if type_name == "Integer" && method == "toString" && !args.is_empty() {
                return java_integer_to_string_call(args);
            }
            if type_name == "BigDecimal" && method == "valueOf" {
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__j_bd_new")),
                    args,
                    optional: false,
                });
            }
            if type_name == "Optional"
                && method == "of"
                && args.len() == 1
                && java_expr_is_long_value(&args[0].value)
            {
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__java_optional_of_long")),
                    args,
                    optional: false,
                });
            }
            if type_name == "NumberFormat" && method == "getCurrencyInstance" {
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__j_df_currency")),
                    args,
                    optional: false,
                });
            }
            if type_name == "NumberFormat" && method == "getPercentInstance" {
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__j_df_percent")),
                    args,
                    optional: false,
                });
            }
            if java_type_simple_name(&type_name) == "MessageFormat" && method == "format" {
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__j_mf_static_format")),
                    args,
                    optional: false,
                });
            }
            if java_type_simple_name(&type_name) == "TimeZone" && method == "getTimeZone" {
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__j_tz_get")),
                    args,
                    optional: false,
                });
            }
            if java_type_simple_name(&type_name) == "Objects" {
                if let Some(expr) = java_objects_static_call(method.as_str(), args.clone()) {
                    return expr;
                }
            }
            if type_name == "StrictMath" && method == "copySign" {
                if java_args_are_copy_sign_negative_zero(&args) {
                    return Expression::string("-0.0");
                }
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__java_double_to_string")),
                    args: vec![Argument::positional(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("StrictMath.copySign")),
                        args,
                        optional: false,
                    }))],
                    optional: false,
                });
            }
            if type_name == "String" && method == "valueOf" {
                let callee = if args.len() == 3 {
                    "__j_array_chars_to_string"
                } else if args.len() == 1 && java_arg_is_char_array(&args[0]) {
                    "__j_string_copy_value_of"
                } else {
                    "__java_string_value_of"
                };
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident(callee)),
                    args,
                    optional: false,
                });
            }
            if type_name == "String" && method == "format" {
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__java_string_format")),
                    args,
                    optional: false,
                });
            }
            if type_name == "String" && method == "join" {
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__java_string_join")),
                    args,
                    optional: false,
                });
            }
            if type_name == "String"
                && matches!(method.as_str(), "copyValueOf" | "translateEscapes")
            {
                let callee = if method == "copyValueOf" && args.len() == 3 {
                    "__j_array_chars_to_string"
                } else if method == "copyValueOf" {
                    "__j_string_copy_value_of"
                } else {
                    "__j_string_translate_escapes"
                };
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident(callee)),
                    args,
                    optional: false,
                });
            }
            if type_name == "Character" {
                if let Some(callee) = java_character_prelude_fn(&method) {
                    return Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident(callee)),
                        args,
                        optional: false,
                    });
                }
            }
            if type_name == "Character.UnicodeBlock" && method == "of" {
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__j_char_unicode_block_of")),
                    args,
                    optional: false,
                });
            }
            if java_type_base_simple_name(&type_name) == "Executors"
                && method == "newFixedThreadPool"
            {
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__j_exec_new")),
                    args,
                    optional: false,
                });
            }
            if java_type_base_simple_name(&type_name) == "Modifier" && args.len() == 1 {
                if let Some(expr) = java_modifier_static_predicate(&method, &args[0].value) {
                    return expr;
                }
            }
            let dotted = format!("{}.{}", type_name, method);
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(&dotted)),
                args,
                optional: false,
            });
        }
    }

    if let ExprKind::Ident(ref type_name) = receiver.kind {
        if java_type_base_simple_name(type_name) == "Modifier" && args.len() == 1 {
            if let Some(expr) = java_modifier_static_predicate(&method, &args[0].value) {
                return expr;
            }
        }
        if type_name == "BigInteger" && method == "valueOf" {
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__java_bigint")),
                args,
                optional: false,
            });
        }
        if type_name == "Executors" && method == "newFixedThreadPool" {
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__j_exec_new")),
                args,
                optional: false,
            });
        }
    }

    if method == "toCharArray" && args.is_empty() {
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__java_to_char_array")),
            args: vec![Argument::positional(receiver)],
            optional: false,
        });
    }

    if method == "formatted" && java_string_method_receiver(&receiver) {
        let mut format_args = Vec::with_capacity(args.len() + 1);
        format_args.push(Argument::positional(receiver));
        format_args.extend(args);
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__java_string_format")),
            args: format_args,
            optional: false,
        });
    }

    if method == "reversed" && args.is_empty() {
        return java_comparator_reversed(receiver);
    }

    if matches!(method.as_str(), "thenComparing" | "thenComparingInt") && args.len() == 1 {
        return java_comparator_then_comparing(receiver, args[0].value.clone());
    }

    if method == "getClass" && args.is_empty() {
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__java_object_get_class")),
            args: vec![Argument::positional(receiver)],
            optional: false,
        });
    }

    if args.is_empty() && java_is_class_token_expr(&receiver) {
        if let Some(expr) = java_class_token_noarg_method(&receiver, method.as_str()) {
            return expr;
        }
        let callee = match method.as_str() {
            "getName" => Some("__java_class_name"),
            "getSimpleName" => Some("__java_class_simple_name"),
            _ => None,
        };
        if let Some(callee) = callee {
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(callee)),
                args: vec![Argument::positional(receiver)],
                optional: false,
            });
        }
    }

    if let Some(expr) = java_class_token_method(&receiver, method.as_str(), &args) {
        return expr;
    }

    if method == "isAssignableFrom" && args.len() == 1 {
        if let Some(expr) = java_class_assignable_from_expr(&receiver, &args[0].value) {
            return expr;
        }
    }

    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(receiver),
            field: method,
            null_safe: false,
        })),
        args,
        optional: false,
    })
}

fn java_reflection_meta(
    class_name: &str,
    parent: Option<String>,
    interfaces: Vec<String>,
    members: &[ClassMember],
    is_interface: bool,
    is_enum: bool,
    class_modifiers: i64,
) -> JavaReflectionClassMeta {
    let mut fields = Vec::new();
    let mut methods = Vec::new();
    let mut constructors = Vec::new();
    let mut nested_classes = Vec::new();

    for member in members {
        match member {
            ClassMember::Field {
                name,
                type_hint,
                modifiers,
                ..
            } => {
                if !name.starts_with("__") {
                    fields.push(JavaReflectionFieldMeta {
                        name: name.clone(),
                        type_name: type_hint.clone(),
                        modifiers: java_member_modifier_bits(modifiers),
                    });
                }
            }
            ClassMember::Method(stmt) => {
                if let StmtKind::FunctionDecl {
                    name,
                    params,
                    return_type,
                    modifiers,
                    ..
                } = &stmt.kind
                {
                    if !name.starts_with("__") {
                        methods.push(JavaReflectionCallableMeta {
                            name: name.clone(),
                            param_count: params.len(),
                            param_types: params
                                .iter()
                                .map(|param| {
                                    param
                                        .type_hint
                                        .clone()
                                        .unwrap_or_else(|| "Object".to_string())
                                })
                                .collect(),
                            return_type: return_type.clone(),
                            modifiers: java_member_modifier_bits(modifiers),
                        });
                    }
                }
            }
            ClassMember::Constructor {
                params, visibility, ..
            } => {
                constructors.push(JavaReflectionCallableMeta {
                    name: class_name.to_string(),
                    param_count: params.len(),
                    param_types: params
                        .iter()
                        .map(|param| {
                            param
                                .type_hint
                                .clone()
                                .unwrap_or_else(|| "Object".to_string())
                        })
                        .collect(),
                    return_type: None,
                    modifiers: java_visibility_modifier_bits(*visibility),
                });
            }
            ClassMember::NestedType(stmt) => {
                if let StmtKind::ClassDecl { name, .. }
                | StmtKind::EnumDecl { name, .. }
                | StmtKind::InterfaceDecl { name, .. } = &stmt.kind
                {
                    nested_classes.push(name.clone());
                }
            }
            _ => {}
        }
    }

    JavaReflectionClassMeta {
        parent,
        interfaces,
        fields,
        methods,
        constructors,
        nested_classes,
        modifiers: class_modifiers,
        is_interface,
        is_enum,
    }
}

fn java_string_method_receiver(receiver: &Expression) -> bool {
    matches!(&receiver.kind, ExprKind::Lit(Literal::Str(_)))
        || matches!(&receiver.kind, ExprKind::Ident(n) if JAVA_STRING_VARS.with(|vars| vars.borrow().contains(n.as_str())))
        || matches!(&receiver.kind, ExprKind::New { class, .. } if matches!(&class.kind, ExprKind::Ident(name) if name == "String"))
        || matches!(&receiver.kind, ExprKind::Call { callee, .. } if matches!(&callee.kind, ExprKind::Ident(name) if name == "__java_string_value_of" || name == "__j_sb_to_string"))
}

fn java_interface_super_receiver(receiver: &Expression) -> Option<String> {
    let ExprKind::Member {
        object,
        field,
        null_safe: false,
    } = &receiver.kind
    else {
        return None;
    };
    if field != "super" {
        return None;
    }
    let ExprKind::Ident(interface_name) = &object.kind else {
        return None;
    };
    JAVA_INTERFACE_NAMES.with(|names| {
        names
            .borrow()
            .contains(interface_name)
            .then(|| interface_name.clone())
    })
}

fn java_reflection_class_meta(name: &str) -> Option<JavaReflectionClassMeta> {
    let simple = java_type_simple_name(name);
    JAVA_REFLECTION_CLASSES.with(|classes| {
        let classes = classes.borrow();
        classes.get(name).or_else(|| classes.get(simple)).cloned()
    })
}

fn java_class_token_method(
    receiver: &Expression,
    method: &str,
    args: &[Argument],
) -> Option<Expression> {
    let class_name = java_class_token_name(receiver)?;
    let meta = java_reflection_class_meta(&class_name);
    match method {
        "getDeclaredFields" | "getFields" if args.is_empty() => Some(java_string_array_expr(
            meta.map(|m| m.fields.into_iter().map(|field| field.name).collect())
                .unwrap_or_default(),
        )),
        "getDeclaredMethods" | "getMethods" if args.is_empty() => {
            Some(java_reflection_callable_array_expr(
                meta.map(|m| m.methods).unwrap_or_default(),
                common_reflection::MEMBER_KIND_METHOD,
                &class_name,
            ))
        }
        "getDeclaredConstructors" | "getConstructors" if args.is_empty() => {
            Some(java_reflection_callable_array_expr(
                meta.map(|m| m.constructors).unwrap_or_default(),
                common_reflection::MEMBER_KIND_CONSTRUCTOR,
                &class_name,
            ))
        }
        "getInterfaces" if args.is_empty() => Some(java_string_array_expr(
            meta.map(|m| m.interfaces).unwrap_or_default(),
        )),
        "getDeclaredClasses" | "getClasses" if args.is_empty() => Some(java_string_array_expr(
            meta.map(|m| m.nested_classes).unwrap_or_default(),
        )),
        "getDeclaredField" | "getField" if args.len() == 1 => {
            let name = java_string_literal(&args[0].value)?;
            let type_name = meta
                .as_ref()
                .and_then(|m| m.fields.iter().find(|field| field.name == name))
                .and_then(|field| field.type_name.clone());
            let modifiers = meta
                .as_ref()
                .and_then(|m| m.fields.iter().find(|field| field.name == name))
                .map(|field| field.modifiers)
                .unwrap_or(JAVA_MOD_PUBLIC);
            Some(java_reflection_token_expr(
                common_reflection::MEMBER_KIND_FIELD,
                &class_name,
                name,
                0,
                type_name,
                None,
                Vec::new(),
                modifiers,
            ))
        }
        "getDeclaredMethod" | "getMethod" if !args.is_empty() => {
            let name = java_string_literal(&args[0].value)?;
            let requested_count = args.len().saturating_sub(1);
            let param_count = meta
                .as_ref()
                .and_then(|m| {
                    m.methods
                        .iter()
                        .find(|method| method.name == name && method.param_count == requested_count)
                })
                .map(|method| method.param_count)
                .unwrap_or(requested_count);
            let method_meta = meta.as_ref().and_then(|m| {
                m.methods
                    .iter()
                    .find(|method| method.name == name && method.param_count == param_count)
            });
            Some(java_reflection_token_expr(
                common_reflection::MEMBER_KIND_METHOD,
                &class_name,
                name,
                param_count,
                None,
                method_meta.and_then(|method| method.return_type.clone()),
                method_meta
                    .map(|method| method.param_types.clone())
                    .unwrap_or_default(),
                method_meta
                    .map(|method| method.modifiers)
                    .unwrap_or(JAVA_MOD_PUBLIC),
            ))
        }
        "getDeclaredConstructor" | "getConstructor" => {
            let requested_count = args.len();
            let param_count = meta
                .as_ref()
                .and_then(|m| {
                    m.constructors
                        .iter()
                        .find(|ctor| ctor.param_count == requested_count)
                })
                .map(|ctor| ctor.param_count)
                .unwrap_or(requested_count);
            let ctor_meta = meta.as_ref().and_then(|m| {
                m.constructors
                    .iter()
                    .find(|ctor| ctor.param_count == param_count)
            });
            Some(java_reflection_token_expr(
                common_reflection::MEMBER_KIND_CONSTRUCTOR,
                &class_name,
                ctor_meta
                    .map(|ctor| ctor.name.as_str())
                    .unwrap_or(java_type_simple_name(&class_name)),
                param_count,
                None,
                None,
                ctor_meta
                    .map(|ctor| ctor.param_types.clone())
                    .unwrap_or_default(),
                ctor_meta
                    .map(|ctor| ctor.modifiers)
                    .unwrap_or(JAVA_MOD_PUBLIC),
            ))
        }
        _ => None,
    }
}

fn java_string_array_expr(values: Vec<String>) -> Expression {
    common_reflection::string_array_expr(values)
}

fn java_reflection_callable_array_expr(
    values: Vec<JavaReflectionCallableMeta>,
    kind: &str,
    owner: &str,
) -> Expression {
    Expression::new(ExprKind::Array(
        values
            .into_iter()
            .map(|value| ArrayElement {
                key: None,
                value: java_reflection_token_expr(
                    kind,
                    owner,
                    &value.name,
                    value.param_count,
                    None,
                    value.return_type,
                    value.param_types,
                    value.modifiers,
                ),
                spread: false,
                by_ref: false,
            })
            .collect(),
    ))
}

fn java_reflection_indexed_token(expr: &Expression) -> Option<Expression> {
    let ExprKind::Index { object, index, .. } = &expr.kind else {
        return None;
    };
    let ExprKind::Array(elems) = &object.kind else {
        return None;
    };
    let ExprKind::Lit(Literal::Int(index)) = &index.kind else {
        return None;
    };
    let index = usize::try_from(*index).ok()?;
    elems.get(index).map(|elem| elem.value.clone())
}

fn java_reflection_token_expr(
    kind: &str,
    owner: &str,
    name: &str,
    param_count: usize,
    type_name: Option<String>,
    return_type: Option<String>,
    param_types: Vec<String>,
    modifiers: i64,
) -> Expression {
    common_reflection::member_token_expr(
        kind,
        owner,
        name,
        param_count,
        type_name,
        return_type,
        param_types,
        modifiers,
    )
}

fn java_reflection_token_name(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Lit(Literal::Str(value)) => Some(value.clone()),
        _ => common_reflection::member_token(expr).map(|token| token.name),
    }
}

fn java_reflection_token(expr: &Expression) -> Option<common_reflection::MemberToken> {
    common_reflection::member_token(expr)
}

fn java_reflection_token_param_count(expr: &Expression) -> Option<usize> {
    java_reflection_token(expr).map(|token| token.param_count)
}

fn java_reflection_token_type_name(expr: &Expression) -> Option<String> {
    java_reflection_token(expr).and_then(|token| token.type_name)
}

fn java_reflection_token_return_type_name(expr: &Expression) -> Option<String> {
    java_reflection_token(expr).and_then(|token| token.return_type)
}

fn java_reflection_token_param_types(expr: &Expression) -> Option<Vec<String>> {
    java_reflection_token(expr).map(|token| token.param_types)
}

fn java_reflection_token_modifiers(expr: &Expression) -> Option<i64> {
    java_reflection_token(expr).map(|token| token.modifiers)
}

fn java_reflection_token_kind(expr: &Expression) -> Option<String> {
    java_reflection_token(expr).map(|token| token.kind)
}

fn java_reflection_token_operation(
    receiver: &Expression,
    method: &str,
    args: &[Argument],
) -> Option<Expression> {
    let token = java_reflection_token(receiver)?;
    match (token.kind.as_str(), method) {
        (common_reflection::MEMBER_KIND_FIELD, "get") if args.len() == 1 => {
            Some(Expression::new(ExprKind::Member {
                object: Box::new(java_reflection_target_expr(&token.owner, &args[0].value)),
                field: token.name,
                null_safe: false,
            }))
        }
        (common_reflection::MEMBER_KIND_FIELD, "set") if args.len() == 2 => {
            Some(Expression::new(ExprKind::Assign {
                target: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(java_reflection_target_expr(&token.owner, &args[0].value)),
                    field: token.name,
                    null_safe: false,
                })),
                value: Box::new(args[1].value.clone()),
            }))
        }
        (common_reflection::MEMBER_KIND_METHOD, "invoke") if !args.is_empty() => {
            let call_args = java_reflection_call_args(&args[1..]);
            Some(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(java_reflection_target_expr(&token.owner, &args[0].value)),
                    field: token.name,
                    null_safe: false,
                })),
                args: call_args.into_iter().map(Argument::positional).collect(),
                optional: false,
            }))
        }
        (common_reflection::MEMBER_KIND_CONSTRUCTOR, "newInstance") => {
            Some(Expression::new(ExprKind::New {
                class: Box::new(Expression::ident(java_type_simple_name(&token.owner))),
                args: java_reflection_call_args(args)
                    .into_iter()
                    .map(Argument::positional)
                    .collect(),
            }))
        }
        _ => None,
    }
}

fn java_reflection_target_expr(owner: &str, target: &Expression) -> Expression {
    if matches!(target.kind, ExprKind::Lit(Literal::Null)) {
        Expression::ident(java_type_simple_name(owner))
    } else {
        target.clone()
    }
}

fn java_reflection_call_args(args: &[Argument]) -> Vec<Expression> {
    if args.len() == 1 {
        if let ExprKind::Array(elems) = &args[0].value.kind {
            return elems.iter().map(|elem| elem.value.clone()).collect();
        }
    }
    args.iter().map(|arg| arg.value.clone()).collect()
}

fn java_class_token_noarg_method(receiver: &Expression, method: &str) -> Option<Expression> {
    let class_name = java_class_token_name(receiver)?;
    match method {
        "getCanonicalName" | "getTypeName" => Some(Expression::string(&class_name)),
        "getPackageName" => Some(Expression::string(&java_class_package_name(&class_name))),
        "getModifiers" => Some(Expression::int(
            java_reflection_class_meta(&class_name)
                .map(|meta| meta.modifiers)
                .unwrap_or(JAVA_MOD_PUBLIC),
        )),
        "isArray" => Some(Expression::bool(class_name.ends_with("[]"))),
        "isPrimitive" => Some(Expression::bool(java_class_name_is_primitive(&class_name))),
        "isInterface" => Some(Expression::bool(java_class_name_is_interface(&class_name))),
        "isEnum" => Some(Expression::bool(java_class_name_is_enum(&class_name))),
        "getComponentType" => Some(
            class_name
                .strip_suffix("[]")
                .map(Expression::string)
                .unwrap_or_else(Expression::null),
        ),
        "getSuperclass" => Some(
            java_class_super_name(&class_name)
                .map(|name| Expression::string(&name))
                .unwrap_or_else(Expression::null),
        ),
        _ => None,
    }
}

fn java_class_package_name(name: &str) -> String {
    let name = name.strip_suffix("[]").unwrap_or(name);
    name.rsplit_once('.')
        .map(|(package, _)| package.to_string())
        .unwrap_or_default()
}

fn java_class_assignable_from_expr(receiver: &Expression, arg: &Expression) -> Option<Expression> {
    let target = java_class_token_name(receiver)?;
    let source = java_class_token_name(arg)?;
    Some(Expression::bool(java_class_is_assignable_from(
        &target, &source,
    )))
}

fn java_class_token_name(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Lit(Literal::Str(value)) => Some(value.clone()),
        ExprKind::Call { callee, args, .. } if matches!(&callee.kind, ExprKind::Ident(name) if name == "__java_object_get_class") => {
            args.first().and_then(|arg| match &arg.value.kind {
                ExprKind::New { class, .. } => java_expr_dotted_name(class),
                _ => None,
            })
        }
        _ => None,
    }
}

fn java_class_name_is_primitive(name: &str) -> bool {
    matches!(
        name,
        "boolean" | "byte" | "char" | "short" | "int" | "long" | "float" | "double" | "void"
    )
}

fn java_class_name_is_interface(name: &str) -> bool {
    let simple = java_type_simple_name(name);
    java_reflection_class_meta(name).is_some_and(|meta| meta.is_interface)
        || matches!(
            simple,
            "List"
                | "Collection"
                | "Set"
                | "Map"
                | "Runnable"
                | "Comparable"
                | "Serializable"
                | "Cloneable"
        )
}

fn java_class_name_is_enum(name: &str) -> bool {
    let simple = java_type_simple_name(name);
    java_reflection_class_meta(name).is_some_and(|meta| meta.is_enum)
        || JAVA_ENUM_VALUES.with(|values| values.borrow().contains_key(simple))
}

fn java_class_super_name(name: &str) -> Option<String> {
    if java_class_name_is_primitive(name)
        || java_class_name_is_interface(name)
        || matches!(java_type_simple_name(name), "Object")
    {
        return None;
    }
    if name.ends_with("[]") {
        return Some("Object".to_string());
    }
    let simple = java_type_simple_name(name);
    java_reflection_class_meta(name)
        .and_then(|meta| meta.parent)
        .or_else(|| {
            if matches!(
                simple,
                "String"
                    | "Integer"
                    | "Long"
                    | "Short"
                    | "Byte"
                    | "Float"
                    | "Double"
                    | "Boolean"
                    | "Character"
                    | "StringBuilder"
            ) {
                Some("Object".to_string())
            } else {
                None
            }
        })
}

fn java_class_is_assignable_from(target: &str, source: &str) -> bool {
    if target == source || java_type_simple_name(target) == java_type_simple_name(source) {
        return true;
    }
    if java_type_simple_name(target) == "Object" && !java_class_name_is_primitive(source) {
        return true;
    }
    let target_simple = java_type_simple_name(target).to_string();
    let mut current = Some(source.to_string());
    while let Some(name) = current {
        if java_type_simple_name(&name) == target_simple {
            return true;
        }
        current = java_class_super_name(&name);
    }
    JAVA_REFLECTION_CLASSES.with(|classes| {
        let classes = classes.borrow();
        classes
            .get(source)
            .or_else(|| classes.get(java_type_simple_name(source)))
            .is_some_and(|meta| {
                meta.interfaces
                    .iter()
                    .any(|iface| java_type_simple_name(iface) == target_simple)
            })
    })
}

fn java_is_class_token_expr(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Lit(Literal::Str(_)) => true,
        ExprKind::Call { callee, .. } => matches!(
            &callee.kind,
            ExprKind::Ident(name)
                if matches!(
                    name.as_str(),
                    "__java_object_get_class"
                        | "__java_class_name"
                        | "__java_class_simple_name"
                        | "__java_enum_set_get_class"
                        | "__j_props_class"
                )
        ),
        _ => false,
    }
}

fn java_expr_dotted_name(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => Some(name.clone()),
        ExprKind::Member { object, field, .. } => {
            let mut prefix = java_expr_dotted_name(object)?;
            prefix.push('.');
            prefix.push_str(field);
            Some(prefix)
        }
        _ => None,
    }
}

fn java_double_static_prelude_fn(method: &str) -> Option<&'static str> {
    match method {
        "compare" => Some("__j_double_compare"),
        "isInfinite" => Some("__j_double_is_infinite"),
        "isFinite" => Some("__j_double_is_finite"),
        _ => None,
    }
}

fn java_base64_static_call(method: &str, args: Vec<Argument>) -> Option<Expression> {
    let callee = match method {
        "getEncoder" => "__j_b64_encoder",
        "getUrlEncoder" => "__j_b64_url_encoder",
        "getMimeEncoder" => "__j_b64_mime_encoder",
        "getDecoder" => "__j_b64_decoder",
        "getUrlDecoder" => "__j_b64_url_decoder",
        "getMimeDecoder" => "__j_b64_mime_decoder",
        _ => return None,
    };
    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(callee)),
        args,
        optional: false,
    }))
}

fn java_system_static_call(method: &str, args: Vec<Argument>) -> Option<Expression> {
    let callee = match method {
        "getProperty" => "__j_system_get_property",
        "setProperty" => "__j_system_set_property",
        "clearProperty" => "__j_system_clear_property",
        "getProperties" => "__j_system_get_properties",
        "getenv" => "__j_system_getenv",
        _ => return None,
    };
    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(callee)),
        args,
        optional: false,
    }))
}

fn java_thread_static_call(method: &str, args: Vec<Argument>) -> Option<Expression> {
    let callee = match method {
        "currentThread" => "__j_thread_current",
        "sleep" => "__j_thread_sleep",
        "interrupted" => "__j_thread_interrupted",
        _ => return None,
    };
    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(callee)),
        args,
        optional: false,
    }))
}

fn java_uri_static_call(method: &str, args: Vec<Argument>) -> Option<Expression> {
    if method != "create" || args.len() != 1 {
        return None;
    }
    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("__j_uri_new")),
        args,
        optional: false,
    }))
}

fn java_args_to_array(args: &[Argument]) -> Expression {
    Expression::new(ExprKind::Array(
        args.iter()
            .map(|arg| ArrayElement {
                key: None,
                value: arg.value.clone(),
                spread: false,
                by_ref: false,
            })
            .collect(),
    ))
}

fn java_boolean_parse_call(args: Vec<Argument>) -> Expression {
    let value = args
        .into_iter()
        .next()
        .map(|arg| arg.value)
        .unwrap_or_else(Expression::null);
    Expression::new(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(value),
        right: Box::new(Expression::string("true")),
    })
}

fn java_integer_to_string_call(args: Vec<Argument>) -> Expression {
    let mut args = args.into_iter();
    let value = args
        .next()
        .unwrap_or_else(|| Argument::positional(Expression::int(0)));
    if let Some(radix) = args.next() {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__j_to_radix")),
            args: vec![value, radix],
            optional: false,
        })
    } else {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__java_string_value_of")),
            args: vec![value],
            optional: false,
        })
    }
}

fn normalise_comparator_static_call(method: &str, args: Vec<Argument>) -> Option<Expression> {
    match method {
        "naturalOrder" if args.is_empty() => Some(java_natural_comparator(false)),
        "reverseOrder" if args.is_empty() => Some(java_natural_comparator(true)),
        "comparing" if args.len() == 1 => Some(java_comparing_comparator(args[0].value.clone())),
        "comparing" if args.len() == 2 => Some(java_comparing_with_comparator(
            args[0].value.clone(),
            args[1].value.clone(),
        )),
        "nullsLast" if args.len() == 1 => Some(java_comparator_nulls_last(args[0].value.clone())),
        _ => None,
    }
}

fn java_lambda_param(name: &str) -> Param {
    Param {
        name: name.to_string(),
        type_hint: None,
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    }
}

fn java_one_arg_lambda(body: Expression) -> Expression {
    Expression::new(ExprKind::Lambda {
        params: vec![java_lambda_param("__value__")],
        body: LambdaBody::Expr(Box::new(body)),
        is_async: false,
        captures: vec![],
    })
}

fn java_two_arg_lambda(body: Expression) -> Expression {
    Expression::new(ExprKind::Lambda {
        params: vec![java_lambda_param("__a__"), java_lambda_param("__b__")],
        body: LambdaBody::Expr(Box::new(body)),
        is_async: false,
        captures: vec![],
    })
}

fn java_functional_value_call(function: Expression, args: Vec<Argument>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Sequence(vec![function]))),
        args,
        optional: false,
    })
}

fn java_binary(op: BinOp, left: Expression, right: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right),
    })
}

fn java_not(expr: Expression) -> Expression {
    Expression::new(ExprKind::Unary {
        op: UnaryOp::Not,
        expr: Box::new(expr),
    })
}

fn java_sequence(items: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Sequence(items))
}

fn java_objects_equals(left: Expression, right: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("__j_objects_equals")),
        args: vec![Argument::positional(left), Argument::positional(right)],
        optional: false,
    })
}

fn java_objects_static_call(method: &str, args: Vec<Argument>) -> Option<Expression> {
    let callee = match method {
        "equals" => "__j_objects_equals",
        "hash" => "__j_objects_hash",
        "hashCode" => "__j_objects_hash_code",
        "requireNonNull" => "__j_objects_require_non_null",
        "isNull" => "__j_objects_is_null",
        "nonNull" => "__j_objects_non_null",
        "compare" => "__j_objects_compare",
        "toString" => "__j_objects_to_string",
        _ => return None,
    };
    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(callee)),
        args,
        optional: false,
    }))
}

fn java_ternary(cond: Expression, then_expr: Expression, else_expr: Expression) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(cond),
        then: Box::new(then_expr),
        else_: Box::new(else_expr),
    })
}

fn java_call(callee: Expression, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(callee),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

fn java_compare_expr(left: Expression, right: Expression, reverse: bool) -> Expression {
    let (less_value, greater_value) = if reverse { (1, -1) } else { (-1, 1) };
    java_ternary(
        java_binary(BinOp::Lt, left.clone(), right.clone()),
        Expression::int(less_value),
        java_ternary(
            java_binary(BinOp::Gt, left, right),
            Expression::int(greater_value),
            Expression::int(0),
        ),
    )
}

fn java_natural_comparator(reverse: bool) -> Expression {
    java_two_arg_lambda(java_compare_expr(
        Expression::ident("__a__"),
        Expression::ident("__b__"),
        reverse,
    ))
}

fn java_key_compare_expr(key_fn: Expression, left_name: &str, right_name: &str) -> Expression {
    let left_key = java_call(key_fn.clone(), vec![Expression::ident(left_name)]);
    let right_key = java_call(key_fn, vec![Expression::ident(right_name)]);
    java_compare_expr(left_key, right_key, false)
}

fn java_comparing_comparator(key_fn: Expression) -> Expression {
    java_two_arg_lambda(java_key_compare_expr(key_fn, "__a__", "__b__"))
}

fn java_comparing_with_comparator(key_fn: Expression, comparator: Expression) -> Expression {
    let left_key = java_call(key_fn.clone(), vec![Expression::ident("__a__")]);
    let right_key = java_call(key_fn, vec![Expression::ident("__b__")]);
    java_two_arg_lambda(java_functional_value_call(
        comparator,
        vec![
            Argument::positional(left_key),
            Argument::positional(right_key),
        ],
    ))
}

fn java_comparator_call(comparator: Expression, left_name: &str, right_name: &str) -> Expression {
    java_call(
        comparator,
        vec![Expression::ident(left_name), Expression::ident(right_name)],
    )
}

fn java_comparator_reversed(comparator: Expression) -> Expression {
    java_two_arg_lambda(java_comparator_call(comparator, "__b__", "__a__"))
}

fn java_comparator_nulls_last(comparator: Expression) -> Expression {
    let a_is_null = java_binary(BinOp::Eq, Expression::ident("__a__"), Expression::null());
    let b_is_null = java_binary(BinOp::Eq, Expression::ident("__b__"), Expression::null());
    java_two_arg_lambda(java_ternary(
        a_is_null,
        java_ternary(b_is_null.clone(), Expression::int(0), Expression::int(1)),
        java_ternary(
            b_is_null,
            Expression::int(-1),
            java_comparator_call(comparator, "__a__", "__b__"),
        ),
    ))
}

fn java_comparator_then_comparing(comparator: Expression, next: Expression) -> Expression {
    let primary_for_cond = java_comparator_call(comparator.clone(), "__a__", "__b__");
    let primary_for_result = java_comparator_call(comparator, "__a__", "__b__");
    let secondary = match &next.kind {
        ExprKind::Lambda { params, .. } if params.len() == 2 => {
            java_comparator_call(next, "__a__", "__b__")
        }
        _ => java_key_compare_expr(next, "__a__", "__b__"),
    };
    java_two_arg_lambda(java_ternary(
        java_binary(BinOp::NotEq, primary_for_cond, Expression::int(0)),
        primary_for_result,
        secondary,
    ))
}

fn java_qualified_static_type(expr: &Expression) -> Option<String> {
    let mut parts = Vec::new();
    collect_member_chain(expr, &mut parts)?;
    if parts.len() < 2 {
        return None;
    }
    let dotted = parts.join(".");
    if JAVA_NESTED_TYPE_NAMES.with(|names| names.borrow().contains(&dotted)) {
        return Some(dotted);
    }
    if !(parts.starts_with(&["java", "util"])
        || parts.starts_with(&["java", "util", "concurrent"])
        || parts.starts_with(&["java", "lang"])
        || parts.starts_with(&["java", "math"])
        || parts.starts_with(&["java", "net"])
        || parts.starts_with(&["java", "nio"])
        || parts.starts_with(&["java", "text"])
        || parts.starts_with(&["java", "time"]))
    {
        return None;
    }
    let type_name = parts.last().copied()?;
    if is_java_type_or_util(type_name) {
        Some(type_name.to_string())
    } else {
        None
    }
}

fn collect_member_chain<'a>(expr: &'a Expression, parts: &mut Vec<&'a str>) -> Option<()> {
    match expr.kind {
        ExprKind::Ident(ref name) => {
            parts.push(name.as_str());
            Some(())
        }
        ExprKind::Member {
            ref object,
            ref field,
            ..
        } => {
            collect_member_chain(object, parts)?;
            parts.push(field.as_str());
            Some(())
        }
        _ => None,
    }
}

fn java_character_prelude_fn(method: &str) -> Option<&'static str> {
    match method {
        "isDigit" => Some("__j_char_is_digit"),
        "isLetter" => Some("__j_char_is_letter"),
        "isLetterOrDigit" => Some("__j_char_is_alnum"),
        "isWhitespace" => Some("__j_char_is_space"),
        "getNumericValue" => Some("__j_char_numeric"),
        "toChars" => Some("__j_char_to_chars"),
        "toCodePoint" => Some("__j_char_to_code_point"),
        "isSurrogate" => Some("__j_char_is_surrogate"),
        "isHighSurrogate" => Some("__j_char_is_high_surrogate"),
        "isLowSurrogate" => Some("__j_char_is_low_surrogate"),
        "highSurrogate" => Some("__j_char_high_surrogate"),
        "lowSurrogate" => Some("__j_char_low_surrogate"),
        "charCount" => Some("__j_char_char_count"),
        "codePointAt" => Some("__j_string_code_point_at"),
        "codePointBefore" => Some("__j_string_code_point_before"),
        "codePointCount" => Some("__j_string_code_point_count"),
        "offsetByCodePoints" => Some("__j_string_offset_by_code_points"),
        "isValidCodePoint" => Some("__j_char_is_valid_code_point"),
        "isBmpCodePoint" => Some("__j_char_is_bmp_code_point"),
        "isSupplementaryCodePoint" => Some("__j_char_is_supplementary_code_point"),
        "digit" => Some("__j_char_digit"),
        "forDigit" => Some("__j_char_for_digit"),
        "compare" => Some("__j_char_compare"),
        "reverseBytes" => Some("__j_char_reverse_bytes"),
        "isDefined" => Some("__j_char_is_defined"),
        "getType" => Some("__j_char_get_type"),
        "isJavaIdentifierStart" => Some("__j_char_is_java_identifier_start"),
        "isJavaIdentifierPart" => Some("__j_char_is_java_identifier_part"),
        "isUnicodeIdentifierStart" => Some("__j_char_is_unicode_identifier_start"),
        "isUnicodeIdentifierPart" => Some("__j_char_is_unicode_identifier_part"),
        "isISOControl" => Some("__j_char_is_iso_control"),
        "isMirrored" => Some("__j_char_is_mirrored"),
        "toTitleCase" => Some("__j_char_to_title_case"),
        "isSurrogatePair" => Some("__j_char_is_surrogate_pair"),
        _ => None,
    }
}

fn java_is_functional_interface_type(type_name: &str) -> bool {
    let simple = java_type_simple_name(type_name);
    matches!(
        simple,
        "Runnable"
            | "Callable"
            | "Function"
            | "BiFunction"
            | "Predicate"
            | "BiPredicate"
            | "Consumer"
            | "BiConsumer"
            | "Supplier"
            | "UnaryOperator"
            | "BinaryOperator"
            | "IntUnaryOperator"
            | "LongUnaryOperator"
            | "DoubleUnaryOperator"
            | "IntPredicate"
            | "LongPredicate"
            | "DoublePredicate"
            | "IntFunction"
            | "LongFunction"
            | "DoubleFunction"
            | "IntConsumer"
            | "LongConsumer"
            | "DoubleConsumer"
            | "IntSupplier"
            | "LongSupplier"
            | "DoubleSupplier"
            | "BooleanSupplier"
    ) || type_name.contains("java.util.function.")
        || type_name.contains("java.util.concurrent.Callable")
        || JAVA_FUNCTIONAL_INTERFACE_METHODS.with(|methods| {
            methods
                .borrow()
                .contains_key(simple.split('<').next().unwrap_or(simple))
        })
}

fn java_functional_receiver(receiver: &Expression) -> bool {
    match &receiver.kind {
        ExprKind::Ident(name) => JAVA_FUNCTIONAL_VARS.with(|vars| vars.borrow().contains(name)),
        ExprKind::Lambda { .. } => true,
        _ => false,
    }
}

fn java_optional_receiver(receiver: &Expression) -> bool {
    match &receiver.kind {
        ExprKind::Ident(name) => JAVA_OPTIONAL_VARS.with(|vars| vars.borrow().contains(name)),
        ExprKind::Call { callee, .. } => {
            matches!(
                &callee.kind,
                ExprKind::Ident(name)
                    if matches!(
                        name.as_str(),
                        "Optional.empty"
                            | "Optional.of"
                            | "Optional.ofNullable"
                            | "__java_optional_filter"
                            | "__java_optional_map"
                            | "__java_optional_flat_map"
                            | "__java_optional_or"
                            | "__java_optional_or_get"
                            | "__java_stream_find_first"
                            | "__java_stream_min"
                            | "__java_stream_max"
                    )
            ) || matches!(
                &callee.kind,
                ExprKind::Member { field, .. } if matches!(field.as_str(), "findFirst" | "findAny" | "min" | "max")
            )
        }
        _ => false,
    }
}

fn java_functional_method(method: &str) -> bool {
    matches!(
        method,
        "apply"
            | "applyAsInt"
            | "applyAsLong"
            | "applyAsDouble"
            | "test"
            | "accept"
            | "get"
            | "getAsInt"
            | "getAsLong"
            | "getAsDouble"
            | "getAsBoolean"
            | "call"
            | "run"
    )
}

fn java_functional_receiver_method(receiver: &Expression, method: &str) -> bool {
    if java_functional_method(method) {
        return true;
    }
    let Some(type_name) = java_functional_type_of(receiver) else {
        return false;
    };
    let simple = java_type_simple_name(&type_name);
    let simple = simple.split('<').next().unwrap_or(simple);
    JAVA_FUNCTIONAL_INTERFACE_METHODS.with(|methods| {
        methods
            .borrow()
            .get(simple)
            .is_some_and(|expected| expected == method)
    })
}

fn java_functional_result_method(method: &str) -> bool {
    matches!(
        method,
        "apply" | "applyAsInt" | "applyAsLong" | "applyAsDouble" | "test" | "accept"
    )
}

fn java_functional_type_of(receiver: &Expression) -> Option<String> {
    match &receiver.kind {
        ExprKind::Ident(name) => {
            JAVA_FUNCTIONAL_TYPES.with(|types| types.borrow().get(name).cloned())
        }
        _ => None,
    }
}

fn java_functional_is_consumer(receiver: &Expression) -> bool {
    java_functional_type_of(receiver)
        .map(|type_name| java_type_simple_name(&type_name).contains("Consumer"))
        .unwrap_or(false)
}

fn java_functional_is_bi(receiver: &Expression) -> bool {
    match &receiver.kind {
        ExprKind::Lambda { params, .. } => params.len() == 2,
        _ => java_functional_type_of(receiver)
            .map(|type_name| java_type_simple_name(&type_name).starts_with("Bi"))
            .unwrap_or(false),
    }
}

fn java_forward_args(is_bi: bool) -> Vec<Argument> {
    if is_bi {
        vec![
            Argument::positional(Expression::ident("__a__")),
            Argument::positional(Expression::ident("__b__")),
        ]
    } else {
        vec![Argument::positional(Expression::ident("__value__"))]
    }
}

fn java_forwarding_lambda(is_bi: bool, body: Expression) -> Expression {
    if is_bi {
        java_two_arg_lambda(body)
    } else {
        java_one_arg_lambda(body)
    }
}

fn java_functional_default_method(
    receiver: &Expression,
    method: &str,
    args: &[Argument],
) -> Option<Expression> {
    let is_bi = java_functional_is_bi(receiver);
    match method {
        "and" if args.len() == 1 => {
            let first = java_functional_value_call(receiver.clone(), java_forward_args(is_bi));
            let second =
                java_functional_value_call(args[0].value.clone(), java_forward_args(is_bi));
            Some(java_forwarding_lambda(
                is_bi,
                java_binary(BinOp::And, first, second),
            ))
        }
        "or" if args.len() == 1 => {
            let first = java_functional_value_call(receiver.clone(), java_forward_args(is_bi));
            let second =
                java_functional_value_call(args[0].value.clone(), java_forward_args(is_bi));
            Some(java_forwarding_lambda(
                is_bi,
                java_binary(BinOp::Or, first, second),
            ))
        }
        "negate" if args.is_empty() => {
            let value = java_functional_value_call(receiver.clone(), java_forward_args(is_bi));
            Some(java_forwarding_lambda(is_bi, java_not(value)))
        }
        "compose" if args.len() == 1 => {
            let before = java_functional_value_call(
                args[0].value.clone(),
                vec![Argument::positional(Expression::ident("__value__"))],
            );
            let after =
                java_functional_value_call(receiver.clone(), vec![Argument::positional(before)]);
            Some(java_one_arg_lambda(after))
        }
        "andThen" if args.len() == 1 && java_functional_is_consumer(receiver) => {
            let first = java_functional_value_call(receiver.clone(), java_forward_args(is_bi));
            let second =
                java_functional_value_call(args[0].value.clone(), java_forward_args(is_bi));
            Some(java_forwarding_lambda(
                is_bi,
                java_sequence(vec![first, second]),
            ))
        }
        "andThen" if args.len() == 1 => {
            let first = java_functional_value_call(receiver.clone(), java_forward_args(is_bi));
            let second = java_functional_value_call(
                args[0].value.clone(),
                vec![Argument::positional(first)],
            );
            Some(java_forwarding_lambda(is_bi, second))
        }
        _ => None,
    }
}

fn java_functional_static_call(
    type_name: &str,
    method: &str,
    args: &[Argument],
) -> Option<Expression> {
    match (java_type_simple_name(type_name), method) {
        ("Function", "identity") | ("UnaryOperator", "identity") if args.is_empty() => {
            Some(java_one_arg_lambda(Expression::ident("__value__")))
        }
        ("Predicate", "isEqual") if args.len() == 1 => Some(java_one_arg_lambda(
            java_objects_equals(Expression::ident("__value__"), args[0].value.clone()),
        )),
        ("BiPredicate", "isEqual") if args.len() == 1 => {
            let first = java_objects_equals(Expression::ident("__a__"), args[0].value.clone());
            let second = java_objects_equals(Expression::ident("__b__"), args[0].value.clone());
            Some(java_two_arg_lambda(java_binary(BinOp::And, first, second)))
        }
        ("BinaryOperator", "minBy") if args.len() == 1 => {
            let compare = java_functional_value_call(
                args[0].value.clone(),
                vec![
                    Argument::positional(Expression::ident("__a__")),
                    Argument::positional(Expression::ident("__b__")),
                ],
            );
            Some(java_two_arg_lambda(java_ternary(
                java_binary(BinOp::LtEq, compare, Expression::int(0)),
                Expression::ident("__a__"),
                Expression::ident("__b__"),
            )))
        }
        ("BinaryOperator", "maxBy") if args.len() == 1 => {
            let compare = java_functional_value_call(
                args[0].value.clone(),
                vec![
                    Argument::positional(Expression::ident("__a__")),
                    Argument::positional(Expression::ident("__b__")),
                ],
            );
            Some(java_two_arg_lambda(java_ternary(
                java_binary(BinOp::GtEq, compare, Expression::int(0)),
                Expression::ident("__a__"),
                Expression::ident("__b__"),
            )))
        }
        _ => None,
    }
}

fn java_arg_is_char_array(arg: &Argument) -> bool {
    matches!(
        &arg.value.kind,
        ExprKind::Ident(name) if JAVA_CHAR_ARRAY_VARS.with(|vars| vars.borrow().contains(name.as_str()))
    )
}

fn java_arg_is_byte_array(arg: &Argument) -> bool {
    matches!(
        &arg.value.kind,
        ExprKind::Ident(name) if JAVA_BYTE_ARRAY_VARS.with(|vars| vars.borrow().contains(name.as_str()))
    )
}

fn java_string_ctor_arg_is_array_source(arg: &Argument) -> bool {
    if java_arg_is_char_array(arg) || java_arg_is_byte_array(arg) {
        return true;
    }
    match &arg.value.kind {
        ExprKind::Array(_) => true,
        ExprKind::Call { callee, .. } => matches!(
            &callee.kind,
            ExprKind::Ident(name)
                if matches!(
                    name.as_str(),
                    "__j_char_to_chars"
                        | "__j_string_get_bytes"
                        | "__j_b64_encode"
                        | "__j_b64_decode"
                )
        ),
        _ => false,
    }
}

fn is_java_type_or_util(name: &str) -> bool {
    matches!(
        name,
        "Integer"
            | "Long"
            | "Short"
            | "Byte"
            | "Float"
            | "Double"
            | "Boolean"
            | "Character"
            | "String"
            | "Math"
            | "StrictMath"
            | "BigDecimal"
            | "BigInteger"
            | "RoundingMode"
            | "DecimalFormat"
            | "DecimalFormatSymbols"
            | "NumberFormat"
            | "MessageFormat"
            | "Locale"
            | "TimeZone"
            | "Calendar"
            | "GregorianCalendar"
            | "URI"
            | "URL"
            | "Arrays"
            | "List"
            | "Set"
            | "Map"
            | "Collections"
            | "Objects"
            | "Optional"
            | "IntStream"
            | "LongStream"
            | "DoubleStream"
            | "Stream"
            | "StreamSupport"
            | "Collectors"
            | "System"
            | "Thread"
            | "ThreadLocalRandom"
            | "Executors"
            | "FutureTask"
            | "Executor"
            | "ExecutorService"
            | "Runtime"
            | "Process"
            | "ProcessBuilder"
            | "File"
            | "Class"
            | "Modifier"
            | "Comparator"
            | "Function"
            | "BiFunction"
            | "Predicate"
            | "BiPredicate"
            | "Consumer"
            | "BiConsumer"
            | "Supplier"
            | "UnaryOperator"
            | "BinaryOperator"
            | "Base64"
            | "Instant"
            | "Duration"
            | "ZoneId"
            | "ZoneOffset"
            | "ChronoUnit"
    )
}

fn walk_primary_atom(pair: Pair<Rule>) -> Result<Expression, String> {
    let inner = pair.into_inner().next().ok_or("primary_atom: empty")?;
    match inner.as_rule() {
        Rule::new_expression => walk_new(inner),
        Rule::array_creation => walk_array_creation(inner),
        Rule::switch_expression => walk_switch_expression(inner),
        Rule::lambda_expression => walk_lambda(inner),
        Rule::paren_expression => walk_expression(inner.into_inner().next().ok_or("paren: empty")?),
        Rule::literal => walk_literal(inner),
        Rule::this_kw => Ok(Expression::new(ExprKind::This)),
        Rule::super_kw => Ok(Expression::new(ExprKind::Super)),
        Rule::super_method_call => walk_super_call(inner),
        Rule::class_literal => Ok(Expression::string(
            inner
                .as_str()
                .strip_suffix(".class")
                .unwrap_or(inner.as_str()),
        )),
        Rule::method_reference => walk_method_reference(inner),
        Rule::ident_name => Ok(Expression::ident(inner.as_str())),
        _ => walk_expression(inner),
    }
}

fn walk_new(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner().peekable();

    let class_name = if inner.peek().map(|p| p.as_rule()) == Some(Rule::type_ref) {
        extract_ref_name(&inner.next().unwrap())
    } else {
        "Object".to_string()
    };
    // skip optional type_args
    if inner.peek().map(|p| p.as_rule()) == Some(Rule::type_args) {
        inner.next();
    }

    let mut anonymous_interfaces = Vec::new();
    while let Some(p) = inner.next() {
        match p.as_rule() {
            Rule::argument_list => {
                let args = walk_arguments(p)?;
                if class_name.rsplit('.').next() == Some("Comparator")
                    && inner.peek().map(|next| next.as_rule()) == Some(Rule::anonymous_class_body)
                {
                    if let Some(comparator) = walk_anonymous_comparator(inner.next().unwrap())? {
                        return Ok(comparator);
                    }
                }
                let mut interfaces = Vec::new();
                while inner.peek().map(|next| next.as_rule()) == Some(Rule::type_ref) {
                    interfaces.push(extract_ref_name(&inner.next().unwrap()));
                }
                if interfaces.is_empty() {
                    interfaces = std::mem::take(&mut anonymous_interfaces);
                }
                if inner.peek().map(|next| next.as_rule()) == Some(Rule::anonymous_class_body) {
                    return walk_anonymous_class_new(
                        &class_name,
                        args,
                        interfaces,
                        inner.next().unwrap(),
                    );
                }
                if matches!(
                    class_name.as_str(),
                    "QName" | "java.xml.namespace.QName" | "javax.xml.namespace.QName"
                ) {
                    let mut values: Vec<Expression> =
                        args.into_iter().map(|arg| arg.value).collect();
                    let call_args = match values.len() {
                        1 => vec![
                            Argument::positional(Expression::string("")),
                            Argument::positional(values.remove(0)),
                            Argument::positional(Expression::string("")),
                        ],
                        2 => vec![
                            Argument::positional(values.remove(0)),
                            Argument::positional(values.remove(0)),
                            Argument::positional(Expression::string("")),
                        ],
                        _ => {
                            let namespace = values
                                .first()
                                .cloned()
                                .unwrap_or_else(|| Expression::string(""));
                            let local = values
                                .get(1)
                                .cloned()
                                .unwrap_or_else(|| Expression::string(""));
                            let prefix = values
                                .get(2)
                                .cloned()
                                .unwrap_or_else(|| Expression::string(""));
                            vec![
                                Argument::positional(namespace),
                                Argument::positional(local),
                                Argument::positional(prefix),
                            ]
                        }
                    };
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__java_xml_name")),
                        args: call_args,
                        optional: false,
                    }));
                }
                // java.net.URI needs to represent opaque and relative strings
                // as well as WHATWG-parsed hierarchical URLs.
                if matches!(class_name.as_str(), "URI" | "java.net.URI") {
                    let ctor = match args.len() {
                        3 => "__j_uri_make3",
                        7 => "__j_uri_make7",
                        _ => "__j_uri_new",
                    };
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident(ctor)),
                        args,
                        optional: false,
                    }));
                }
                if matches!(
                    class_name.as_str(),
                    "ProcessBuilder" | "java.lang.ProcessBuilder"
                ) {
                    let args = vec![Argument::positional(java_args_to_array(&args))];
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__j_pb_new")),
                        args,
                        optional: false,
                    }));
                }
                if matches!(class_name.as_str(), "File" | "java.io.File") {
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__j_file_new")),
                        args,
                        optional: false,
                    }));
                }
                // java.net.URL → the WHATWG-parsed object (web:url) the
                // __j_url_* prelude getters read. Arity picks the java
                // constructor form: (spec), (context, spec), or
                // (protocol, host, port, file).
                if matches!(class_name.as_str(), "URL" | "java.net.URL") {
                    let (ctor, args) = match args.len() {
                        2 => ("__j_url_ctx", args),
                        3 => {
                            // URL(protocol, host, file) == port -1 form.
                            let mut args = args;
                            args.insert(2, Argument::positional(Expression::int(-1)));
                            ("__j_url_make", args)
                        }
                        4 => ("__j_url_make", args),
                        _ => ("__j_url_new", args),
                    };
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident(ctor)),
                        args,
                        optional: false,
                    }));
                }
                if matches!(class_name.as_str(), "Properties" | "java.util.Properties") {
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__j_props_new")),
                        args,
                        optional: false,
                    }));
                }
                if matches!(java_type_simple_name(&class_name), "BigDecimal") {
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__j_bd_new")),
                        args,
                        optional: false,
                    }));
                }
                if matches!(java_type_simple_name(&class_name), "DecimalFormatSymbols") {
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__j_df_symbols")),
                        args,
                        optional: false,
                    }));
                }
                if matches!(java_type_simple_name(&class_name), "DecimalFormat") {
                    let mut args = args;
                    if args.len() == 1 {
                        args.push(Argument::positional(Expression::new(ExprKind::Lit(
                            Literal::Undefined,
                        ))));
                    }
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__j_df_new")),
                        args,
                        optional: false,
                    }));
                }
                if matches!(java_type_simple_name(&class_name), "Formatter") {
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__j_fmt_new")),
                        args,
                        optional: false,
                    }));
                }
                if matches!(java_type_simple_name(&class_name), "MessageFormat") {
                    let mut args = args;
                    if args.len() == 1 {
                        args.push(Argument::positional(Expression::string("US")));
                    }
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__j_mf_new")),
                        args,
                        optional: false,
                    }));
                }
                if matches!(
                    java_type_simple_name(&class_name),
                    "GregorianCalendar" | "Calendar"
                ) {
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__j_cal_new")),
                        args,
                        optional: false,
                    }));
                }
                if matches!(
                    class_name.as_str(),
                    "StringBuilder"
                        | "StringBuffer"
                        | "java.lang.StringBuilder"
                        | "java.lang.StringBuffer"
                ) {
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__j_sb_new")),
                        args,
                        optional: false,
                    }));
                }
                if matches!(java_type_simple_name(&class_name), "String") && args.len() == 3 {
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__j_code_points_to_string")),
                        args,
                        optional: false,
                    }));
                }
                if matches!(java_type_simple_name(&class_name), "String") && args.len() == 1 {
                    if java_string_ctor_arg_is_array_source(&args[0]) {
                        return Ok(Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident("__j_string_from_array")),
                            args,
                            optional: false,
                        }));
                    }
                    return Ok(Expression::new(ExprKind::New {
                        class: Box::new(Expression::ident("String")),
                        args,
                    }));
                }
                if matches!(
                    class_name.as_str(),
                    "StringJoiner" | "java.util.StringJoiner"
                ) {
                    let callee = if args.len() == 3 {
                        "__j_sj_new3"
                    } else {
                        "__j_sj_new"
                    };
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident(callee)),
                        args,
                        optional: false,
                    }));
                }
                if matches!(
                    class_name.as_str(),
                    "StringTokenizer" | "java.util.StringTokenizer"
                ) {
                    let callee = match args.len() {
                        2 => "__j_st_new2",
                        3 => "__j_st_new3",
                        _ => "__j_st_new",
                    };
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident(callee)),
                        args,
                        optional: false,
                    }));
                }
                if matches!(class_name.as_str(), "Scanner" | "java.util.Scanner") {
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__j_sc_new")),
                        args,
                        optional: false,
                    }));
                }
                if matches!(class_name.as_str(), "Thread" | "java.lang.Thread") {
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__j_thread_new")),
                        args,
                        optional: false,
                    }));
                }
                if matches!(
                    class_name.as_str(),
                    "Semaphore" | "java.util.concurrent.Semaphore"
                ) {
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__java_semaphore_new")),
                        args,
                        optional: false,
                    }));
                }
                if matches!(java_type_base_simple_name(&class_name), "CountDownLatch") {
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__j_latch_new")),
                        args,
                        optional: false,
                    }));
                }
                if matches!(java_type_base_simple_name(&class_name), "FutureTask") {
                    let mut args = args;
                    let has_preset = args.len() >= 2;
                    if args.len() == 1 {
                        args.push(Argument::positional(Expression::new(ExprKind::Lit(
                            Literal::Undefined,
                        ))));
                    }
                    args.push(Argument::positional(Expression::bool(has_preset)));
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__j_future_task_new")),
                        args,
                        optional: false,
                    }));
                }
                return Ok(Expression::new(ExprKind::New {
                    class: Box::new(Expression::ident(&class_name)),
                    args,
                }));
            }
            Rule::type_ref => {
                anonymous_interfaces.push(extract_ref_name(&p));
            }
            Rule::array_initializer => {
                // new Type[] {1, 2, 3} → array literal
                return walk_initializer_as_array(p);
            }
            Rule::array_dims => {
                let mut sizes = Vec::new();
                let mut initializer = None;
                for dim in p.into_inner() {
                    match dim.as_rule() {
                        Rule::expression => {
                            if let Ok(size) = walk_expression(dim) {
                                sizes.push(size);
                            }
                        }
                        Rule::array_initializer => initializer = Some(dim),
                        _ => {}
                    }
                }
                if sizes.is_empty() {
                    if let Some(init) = initializer {
                        return walk_initializer_as_array(init);
                    }
                }
                if sizes.len() >= 2
                    && matches!(
                        class_name.as_str(),
                        "byte"
                            | "short"
                            | "int"
                            | "long"
                            | "char"
                            | "byte[]"
                            | "short[]"
                            | "int[]"
                            | "long[]"
                            | "char[]"
                    )
                {
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__new_int_2d_array")),
                        args: vec![
                            Argument::positional(sizes[0].clone()),
                            Argument::positional(sizes[1].clone()),
                        ],
                        optional: false,
                    }));
                }
                // new int[5] → __new_array(5)
                if let Some(sz) = sizes.into_iter().next() {
                    let callee = match class_name.as_str() {
                        "boolean" | "boolean[]" => "__new_bool_array",
                        "byte" | "short" | "int" | "long" | "char" | "byte[]" | "short[]"
                        | "int[]" | "long[]" | "char[]" => "__new_int_array",
                        _ => "__new_array",
                    };
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident(callee)),
                        args: vec![Argument::positional(sz)],
                        optional: false,
                    }));
                }
            }
            Rule::anonymous_class_body => {
                if class_name.contains("Comparator") {
                    if let Some(comparator) = walk_anonymous_comparator(p.clone())? {
                        return Ok(comparator);
                    }
                }
                return walk_anonymous_class_new(
                    &class_name,
                    vec![],
                    std::mem::take(&mut anonymous_interfaces),
                    p,
                );
            }
            _ => {}
        }
    }

    if matches!(java_type_simple_name(&class_name), "Formatter") {
        return Ok(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__j_fmt_new")),
            args: vec![],
            optional: false,
        }));
    }

    Ok(Expression::new(ExprKind::New {
        class: Box::new(Expression::ident(&class_name)),
        args: vec![],
    }))
}

fn walk_anonymous_class_new(
    class_name: &str,
    args: Vec<Argument>,
    mut interfaces: Vec<String>,
    body: Pair<Rule>,
) -> Result<Expression, String> {
    let members = walk_class_body(body)?;
    let parent = if interfaces.is_empty() && java_anonymous_interface_target(class_name) {
        interfaces.push(class_name.to_string());
        None
    } else if interfaces.is_empty() && java_anonymous_root_class(class_name) {
        None
    } else {
        Some(Box::new(Expression::ident(class_name)))
    };
    let class_expr = Expression::new(ExprKind::ClassExpr {
        name: None,
        parent,
        interfaces,
        members,
    });
    Ok(Expression::new(ExprKind::New {
        class: Box::new(class_expr),
        args,
    }))
}

fn java_anonymous_interface_target(class_name: &str) -> bool {
    let simple_name = class_name.rsplit('.').next().unwrap_or(class_name);
    matches!(simple_name, "Runnable")
        || JAVA_INTERFACE_NAMES.with(|names| names.borrow().contains(simple_name))
}

fn java_anonymous_root_class(class_name: &str) -> bool {
    matches!(
        class_name.rsplit('.').next().unwrap_or(class_name),
        "Object"
    )
}

fn erase_java_interface_param_hints(body: &mut [Statement]) {
    erase_java_interface_param_hints_with_types(body, &mut HashMap::new());
}

fn erase_java_interface_param_hints_with_types(
    body: &mut [Statement],
    concrete_locals: &mut HashMap<String, String>,
) {
    for stmt in body {
        match &mut stmt.kind {
            StmtKind::FunctionDecl {
                params,
                return_type,
                body,
                ..
            } => {
                erase_java_interface_params(params);
                if return_type
                    .as_deref()
                    .is_some_and(java_anonymous_interface_target)
                {
                    *return_type = None;
                }
                erase_java_interface_param_hints_with_types(body, &mut HashMap::new());
            }
            StmtKind::VarDecl { declarations, .. } => {
                for decl in declarations {
                    if let Some(init) = &mut decl.init {
                        erase_java_interface_hints_expr(init);
                    }
                    let inferred_type = decl
                        .init
                        .as_ref()
                        .and_then(|init| java_concrete_initializer_type(init, concrete_locals));
                    let is_interface_hint = decl
                        .type_hint
                        .as_deref()
                        .is_some_and(java_anonymous_interface_target);
                    if is_interface_hint {
                        decl.type_hint = inferred_type.clone();
                    }
                    if let BindingPattern::Ident(name) = &decl.pattern {
                        if is_interface_hint
                            && decl.init.as_ref().is_some_and(|init| {
                                java_initializer_is_class_object(init, concrete_locals)
                            })
                        {
                            JAVA_FUNCTIONAL_VARS.with(|vars| {
                                vars.borrow_mut().remove(name);
                            });
                            JAVA_FUNCTIONAL_TYPES.with(|types| {
                                types.borrow_mut().remove(name);
                            });
                        }
                        if let Some(type_name) = inferred_type.or_else(|| decl.type_hint.clone()) {
                            concrete_locals.insert(name.clone(), type_name);
                        }
                    }
                }
            }
            StmtKind::Assign { targets, value } => {
                for target in targets {
                    erase_java_interface_hints_expr(target);
                }
                erase_java_interface_hints_expr(value);
            }
            StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
                erase_java_interface_hints_expr(expr);
            }
            StmtKind::ClassDecl { members, .. }
            | StmtKind::StructDecl { members, .. }
            | StmtKind::EnumDecl {
                body_members: members,
                ..
            } => {
                for member in members {
                    match member {
                        ClassMember::Constructor { params, body, .. } => {
                            erase_java_interface_params(params);
                            erase_java_interface_param_hints_with_types(body, &mut HashMap::new());
                        }
                        ClassMember::Method(method) => {
                            erase_java_interface_param_hints_with_types(
                                std::slice::from_mut(method),
                                &mut HashMap::new(),
                            );
                        }
                        ClassMember::NestedType(nested) => {
                            erase_java_interface_param_hints_with_types(
                                std::slice::from_mut(nested),
                                &mut HashMap::new(),
                            );
                        }
                        _ => {}
                    }
                }
            }
            StmtKind::Block(stmts) | StmtKind::NamespaceDecl { body: stmts, .. } => {
                erase_java_interface_param_hints_with_types(stmts, &mut concrete_locals.clone());
            }
            _ => {}
        }
    }
}

fn java_initializer_is_class_object(
    expr: &Expression,
    concrete_locals: &HashMap<String, String>,
) -> bool {
    match &expr.kind {
        ExprKind::New { .. } => true,
        ExprKind::Cast { expr, .. } => java_initializer_is_class_object(expr, concrete_locals),
        ExprKind::Ident(name) => concrete_locals.contains_key(name),
        ExprKind::Sequence(items) => items
            .last()
            .is_some_and(|expr| java_initializer_is_class_object(expr, concrete_locals)),
        _ => false,
    }
}

fn java_concrete_initializer_type(
    expr: &Expression,
    concrete_locals: &HashMap<String, String>,
) -> Option<String> {
    match &expr.kind {
        ExprKind::New { class, .. } => match &class.kind {
            ExprKind::Ident(name) => Some(name.clone()),
            _ => None,
        },
        ExprKind::Cast { expr, .. } => java_concrete_initializer_type(expr, concrete_locals),
        ExprKind::Ident(name) => concrete_locals.get(name).cloned(),
        ExprKind::Sequence(items) => items
            .last()
            .and_then(|expr| java_concrete_initializer_type(expr, concrete_locals)),
        _ => None,
    }
}

fn erase_java_interface_hints_expr(expr: &mut Expression) {
    match &mut expr.kind {
        ExprKind::Cast {
            expr: inner,
            type_name,
        } if java_anonymous_interface_target(type_name) => {
            *expr = (**inner).clone();
        }
        ExprKind::Call { callee, args, .. } => {
            erase_java_interface_hints_expr(callee);
            for arg in args {
                erase_java_interface_hints_expr(&mut arg.value);
            }
        }
        ExprKind::Member { object, .. } => erase_java_interface_hints_expr(object),
        ExprKind::Index { object, index, .. } => {
            erase_java_interface_hints_expr(object);
            erase_java_interface_hints_expr(index);
        }
        ExprKind::Binary { left, right, .. } => {
            erase_java_interface_hints_expr(left);
            erase_java_interface_hints_expr(right);
        }
        ExprKind::Unary { expr, .. } | ExprKind::Cast { expr, .. } => {
            erase_java_interface_hints_expr(expr);
        }
        ExprKind::Assign { target, value } => {
            erase_java_interface_hints_expr(target);
            erase_java_interface_hints_expr(value);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            erase_java_interface_hints_expr(cond);
            erase_java_interface_hints_expr(then);
            erase_java_interface_hints_expr(else_);
        }
        ExprKind::Array(elems) => {
            for elem in elems {
                erase_java_interface_hints_expr(&mut elem.value);
            }
        }
        ExprKind::New { class, args } => {
            erase_java_interface_hints_expr(class);
            for arg in args {
                erase_java_interface_hints_expr(&mut arg.value);
            }
        }
        ExprKind::StaticAccess { class, member } => {
            erase_java_interface_hints_expr(class);
            erase_java_interface_hints_expr(member);
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(expr) => erase_java_interface_hints_expr(expr),
            LambdaBody::Block(stmts) => erase_java_interface_param_hints(stmts),
        },
        _ => {}
    }
}

fn erase_java_interface_params(params: &mut [Param]) {
    for param in params {
        if param
            .type_hint
            .as_deref()
            .is_some_and(java_anonymous_interface_target)
        {
            param.type_hint = None;
        }
    }
}

fn strip_java_abstract_method_declarations(body: &mut [Statement]) {
    for stmt in body {
        match &mut stmt.kind {
            StmtKind::ClassDecl { members, .. }
            | StmtKind::StructDecl { members, .. }
            | StmtKind::EnumDecl {
                body_members: members,
                ..
            } => {
                members.retain(|member| !java_is_abstract_method_declaration(member));
                for member in members {
                    if let ClassMember::NestedType(nested) = member {
                        strip_java_abstract_method_declarations(std::slice::from_mut(nested));
                    }
                }
            }
            StmtKind::Block(stmts) | StmtKind::NamespaceDecl { body: stmts, .. } => {
                strip_java_abstract_method_declarations(stmts);
            }
            _ => {}
        }
    }
}

fn java_is_abstract_method_declaration(member: &ClassMember) -> bool {
    let ClassMember::Method(method) = member else {
        return false;
    };
    let StmtKind::FunctionDecl {
        body, modifiers, ..
    } = &method.kind
    else {
        return false;
    };
    body.is_empty() && !modifiers.is_static
}

fn lower_java_abstract_runtime_modifiers(body: &mut [Statement]) {
    for stmt in body {
        match &mut stmt.kind {
            StmtKind::ClassDecl {
                members, modifiers, ..
            } => {
                modifiers.is_abstract = false;
                for member in members {
                    if let ClassMember::NestedType(nested) = member {
                        lower_java_abstract_runtime_modifiers(std::slice::from_mut(nested));
                    }
                }
            }
            StmtKind::StructDecl { members, .. } => {
                for member in members {
                    if let ClassMember::NestedType(nested) = member {
                        lower_java_abstract_runtime_modifiers(std::slice::from_mut(nested));
                    }
                }
            }
            StmtKind::EnumDecl { body_members, .. } => {
                for member in body_members {
                    if let ClassMember::NestedType(nested) = member {
                        lower_java_abstract_runtime_modifiers(std::slice::from_mut(nested));
                    }
                }
            }
            StmtKind::Block(stmts) | StmtKind::NamespaceDecl { body: stmts, .. } => {
                lower_java_abstract_runtime_modifiers(stmts);
            }
            _ => {}
        }
    }
}

fn reject_java_direct_abstract_instantiation(body: &[Statement]) -> Result<(), String> {
    let mut abstract_classes = HashSet::new();
    collect_java_abstract_classes(body, &mut abstract_classes);
    for stmt in body {
        reject_java_direct_abstract_instantiation_stmt(stmt, &abstract_classes)?;
    }
    Ok(())
}

fn collect_java_abstract_classes(body: &[Statement], out: &mut HashSet<String>) {
    for stmt in body {
        match &stmt.kind {
            StmtKind::ClassDecl {
                name,
                members,
                modifiers,
                ..
            } => {
                if modifiers.is_abstract {
                    out.insert(name.clone());
                }
                for member in members {
                    if let ClassMember::NestedType(nested) = member {
                        collect_java_abstract_classes(std::slice::from_ref(nested), out);
                    }
                }
            }
            StmtKind::Block(stmts) | StmtKind::NamespaceDecl { body: stmts, .. } => {
                collect_java_abstract_classes(stmts, out);
            }
            _ => {}
        }
    }
}

fn reject_java_direct_abstract_instantiation_stmt(
    stmt: &Statement,
    abstract_classes: &HashSet<String>,
) -> Result<(), String> {
    match &stmt.kind {
        StmtKind::Expr(expr)
        | StmtKind::Return(Some(expr))
        | StmtKind::Throw {
            expr: Some(expr), ..
        } => {
            reject_java_direct_abstract_instantiation_expr(expr, abstract_classes)?;
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &decl.init {
                    reject_java_direct_abstract_instantiation_expr(init, abstract_classes)?;
                }
            }
        }
        StmtKind::FunctionDecl { body, .. } | StmtKind::Block(body) => {
            for stmt in body {
                reject_java_direct_abstract_instantiation_stmt(stmt, abstract_classes)?;
            }
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::EnumDecl {
            body_members: members,
            ..
        } => {
            for member in members {
                match member {
                    ClassMember::Field {
                        init: Some(init), ..
                    } => {
                        reject_java_direct_abstract_instantiation_expr(init, abstract_classes)?;
                    }
                    ClassMember::Constructor { body, .. } => {
                        for stmt in body {
                            reject_java_direct_abstract_instantiation_stmt(stmt, abstract_classes)?;
                        }
                    }
                    ClassMember::Method(method) | ClassMember::NestedType(method) => {
                        reject_java_direct_abstract_instantiation_stmt(method, abstract_classes)?;
                    }
                    _ => {}
                }
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            reject_java_direct_abstract_instantiation_expr(cond, abstract_classes)?;
            for stmt in then_body {
                reject_java_direct_abstract_instantiation_stmt(stmt, abstract_classes)?;
            }
            for (cond, body) in elifs {
                reject_java_direct_abstract_instantiation_expr(cond, abstract_classes)?;
                for stmt in body {
                    reject_java_direct_abstract_instantiation_stmt(stmt, abstract_classes)?;
                }
            }
            if let Some(else_body) = else_body {
                for stmt in else_body {
                    reject_java_direct_abstract_instantiation_stmt(stmt, abstract_classes)?;
                }
            }
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                reject_java_direct_abstract_instantiation_expr(target, abstract_classes)?;
            }
            reject_java_direct_abstract_instantiation_expr(value, abstract_classes)?;
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            reject_java_direct_abstract_instantiation_expr(target, abstract_classes)?;
            reject_java_direct_abstract_instantiation_expr(value, abstract_classes)?;
        }
        _ => {}
    }
    Ok(())
}

fn reject_java_direct_abstract_instantiation_expr(
    expr: &Expression,
    abstract_classes: &HashSet<String>,
) -> Result<(), String> {
    match &expr.kind {
        ExprKind::New { class, args } => {
            if let ExprKind::Ident(name) = &class.kind {
                if abstract_classes.contains(name) {
                    return Err(format!("cannot instantiate abstract class {name}"));
                }
            }
            reject_java_direct_abstract_instantiation_expr(class, abstract_classes)?;
            for arg in args {
                reject_java_direct_abstract_instantiation_expr(&arg.value, abstract_classes)?;
            }
        }
        ExprKind::Binary { left, right, .. } => {
            reject_java_direct_abstract_instantiation_expr(left, abstract_classes)?;
            reject_java_direct_abstract_instantiation_expr(right, abstract_classes)?;
        }
        ExprKind::Unary { expr, .. } => {
            reject_java_direct_abstract_instantiation_expr(expr, abstract_classes)?;
        }
        ExprKind::Member { object, .. } => {
            reject_java_direct_abstract_instantiation_expr(object, abstract_classes)?;
        }
        ExprKind::Call { callee, args, .. } => {
            reject_java_direct_abstract_instantiation_expr(callee, abstract_classes)?;
            for arg in args {
                reject_java_direct_abstract_instantiation_expr(&arg.value, abstract_classes)?;
            }
        }
        ExprKind::Assign { target, value } => {
            reject_java_direct_abstract_instantiation_expr(target, abstract_classes)?;
            reject_java_direct_abstract_instantiation_expr(value, abstract_classes)?;
        }
        ExprKind::Array(elems) => {
            for elem in elems {
                reject_java_direct_abstract_instantiation_expr(&elem.value, abstract_classes)?;
            }
        }
        _ => {}
    }
    Ok(())
}

fn walk_anonymous_comparator(pair: Pair<Rule>) -> Result<Option<Expression>, String> {
    for member in pair.into_inner() {
        if member.as_rule() != Rule::method_declaration {
            continue;
        }
        let ClassMember::Method(method) = walk_method(member)? else {
            continue;
        };
        let StmtKind::FunctionDecl {
            name, params, body, ..
        } = method.kind
        else {
            continue;
        };
        if name == "compare" && params.len() == 2 {
            return Ok(Some(Expression::new(ExprKind::Lambda {
                params,
                body: LambdaBody::Block(body),
                is_async: false,
                captures: vec![],
            })));
        }
    }
    Ok(None)
}

fn walk_array_creation(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner();
    let prim_type = inner
        .next()
        .map(|p| p.as_str().to_string())
        .unwrap_or_else(|| "Object".to_string());
    for p in inner {
        match p.as_rule() {
            Rule::array_initializer => return walk_initializer_as_array(p),
            Rule::expression => {
                let sz = walk_expression(p)?;
                let callee = match prim_type.as_str() {
                    "boolean" => "__new_bool_array",
                    "byte" | "short" | "int" | "long" | "char" => "__new_int_array",
                    _ => "__new_array",
                };
                return Ok(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident(callee)),
                    args: vec![Argument::positional(sz)],
                    optional: false,
                }));
            }
            _ => {}
        }
    }
    Ok(Expression::new(ExprKind::Array(vec![])))
}

fn walk_initializer_as_array(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut elems = Vec::new();
    for el in pair.into_inner() {
        if el.as_rule() == Rule::initializer {
            elems.push(ArrayElement {
                key: None,
                value: walk_initializer(el)?,
                spread: false,
                by_ref: false,
            });
        }
    }
    Ok(Expression::new(ExprKind::Array(elems)))
}

fn walk_super_call(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner();
    let method_name = inner
        .next()
        .ok_or("super call: missing name")?
        .as_str()
        .to_string();
    let args = if let Some(al) = inner.next() {
        if al.as_rule() == Rule::argument_list {
            walk_arguments(al)?
        } else {
            vec![]
        }
    } else {
        vec![]
    };
    Ok(Expression::new(ExprKind::SuperCall {
        method: Some(method_name),
        args,
    }))
}

fn walk_method_reference(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner();
    let obj = inner.next().ok_or("method ref: missing object")?;
    let obj_name = obj.as_str().to_string();
    let method = inner
        .next()
        .ok_or("method ref: missing method")?
        .as_str()
        .to_string();

    let obj_expr = Expression::ident(&obj_name);
    if method == "new" {
        if let Some(element_type) = obj_name.strip_suffix("[]") {
            let callee = match element_type {
                "boolean" => "__new_bool_array",
                "byte" | "short" | "int" | "long" | "char" => "__new_int_array",
                _ => "__new_array",
            };
            return Ok(Expression::new(ExprKind::Lambda {
                params: vec![Param {
                    name: "__size__".to_string(),
                    type_hint: None,
                    default: None,
                    pass_by: PassBy::Value,
                    is_rest: false,
                    is_kwargs: false,
                    is_optional: false,
                    is_nullable: false,
                }],
                body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident(callee)),
                    args: vec![Argument::positional(Expression::ident("__size__"))],
                    optional: false,
                }))),
                is_async: false,
                captures: vec![],
            }));
        }
        return Ok(Expression::new(ExprKind::Lambda {
            params: vec![Param {
                name: "__args__".to_string(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: true,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            }],
            body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::New {
                class: Box::new(obj_expr),
                args: vec![],
            }))),
            is_async: false,
            captures: vec![],
        }));
    }

    if matches!(
        (obj_name.as_str(), method.as_str()),
        ("Integer", "max")
            | ("Integer", "min")
            | ("Long", "max")
            | ("Long", "min")
            | ("Double", "max")
            | ("Double", "min")
            | ("Math", "max")
            | ("Math", "min")
            | ("StrictMath", "max")
            | ("StrictMath", "min")
    ) {
        let callee = format!("{}.{}", obj_name, method);
        return Ok(java_two_arg_lambda(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident(&callee)),
            args: vec![
                Argument::positional(Expression::ident("__a__")),
                Argument::positional(Expression::ident("__b__")),
            ],
            optional: false,
        })));
    }

    if matches!(
        (obj_name.as_str(), method.as_str()),
        ("Integer", "sum") | ("Long", "sum")
    ) {
        return Ok(java_two_arg_lambda(Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(Expression::ident("__a__")),
            right: Box::new(Expression::ident("__b__")),
        })));
    }

    if obj_name == "String" && method == "concat" {
        return Ok(java_two_arg_lambda(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__java_string_concat")),
            args: vec![
                Argument::positional(Expression::ident("__a__")),
                Argument::positional(Expression::ident("__b__")),
            ],
            optional: false,
        })));
    }

    if matches!(obj_name.as_str(), "Objects" | "java.util.Objects") && method == "requireNonNull" {
        return Ok(Expression::new(ExprKind::Lambda {
            params: vec![Param {
                name: "__value__".to_string(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            }],
            body: LambdaBody::Expr(Box::new(Expression::ident("__value__"))),
            is_async: false,
            captures: vec![],
        }));
    }

    if matches!(obj_name.as_str(), "Optional" | "java.util.Optional") && method == "empty" {
        return Ok(Expression::new(ExprKind::Lambda {
            params: vec![],
            body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("Optional.empty")),
                args: vec![],
                optional: false,
            }))),
            is_async: false,
            captures: vec![],
        }));
    }

    if matches!(obj_name.as_str(), "Optional" | "java.util.Optional") && method == "isPresent" {
        return Ok(Expression::new(ExprKind::Lambda {
            params: vec![Param {
                name: "__value__".to_string(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            }],
            body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__java_optional_is_present")),
                args: vec![Argument::positional(Expression::ident("__value__"))],
                optional: false,
            }))),
            is_async: false,
            captures: vec![],
        }));
    }

    if obj_name == "Math" || obj_name == "StrictMath" {
        let callee = format!("{}.{}", obj_name, method);
        return Ok(Expression::new(ExprKind::Lambda {
            params: vec![Param {
                name: "__value__".to_string(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            }],
            body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(&callee)),
                args: vec![Argument::positional(Expression::ident("__value__"))],
                optional: false,
            }))),
            is_async: false,
            captures: vec![],
        }));
    }

    if obj_name == "System.out" && matches!(method.as_str(), "print" | "println") {
        let callee = if method == "println" {
            "__java_println"
        } else {
            "__java_print"
        };
        return Ok(Expression::new(ExprKind::Lambda {
            params: vec![Param {
                name: "__value__".to_string(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            }],
            body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(callee)),
                args: vec![Argument::positional(Expression::ident("__value__"))],
                optional: false,
            }))),
            is_async: false,
            captures: vec![],
        }));
    }

    if matches!(
        (obj_name.as_str(), method.as_str()),
        ("Integer", "parseInt")
            | ("Integer", "valueOf")
            | ("Long", "parseLong")
            | ("Double", "parseDouble")
            | ("String", "valueOf")
    ) {
        let callee = format!("{}.{}", obj_name, method);
        return Ok(Expression::new(ExprKind::Lambda {
            params: vec![Param {
                name: "__value__".to_string(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            }],
            body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(
                    if obj_name == "String" && method == "valueOf" {
                        "__java_string_value_of"
                    } else {
                        &callee
                    },
                )),
                args: vec![Argument::positional(Expression::ident("__value__"))],
                optional: false,
            }))),
            is_async: false,
            captures: vec![],
        }));
    }

    if obj_name == "String" && method == "compareTo" {
        return Ok(java_two_arg_lambda(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__j_string_compare_to")),
            args: vec![
                Argument::positional(Expression::ident("__a__")),
                Argument::positional(Expression::ident("__b__")),
            ],
            optional: false,
        })));
    }

    if matches!(
        (obj_name.as_str(), method.as_str()),
        ("Integer", "compareTo")
            | ("Long", "compareTo")
            | ("Double", "compareTo")
            | ("Float", "compareTo")
            | ("Short", "compareTo")
            | ("Byte", "compareTo")
    ) {
        return Ok(java_two_arg_lambda(java_compare_expr(
            Expression::ident("__a__"),
            Expression::ident("__b__"),
            false,
        )));
    }

    if matches!(
        (obj_name.as_str(), method.as_str()),
        ("String", "length")
            | ("String", "toString")
            | ("String", "strip")
            | ("String", "trim")
            | ("String", "toUpperCase")
            | ("String", "toLowerCase")
            | ("String", "isEmpty")
            | ("String", "isBlank")
            | ("Integer", "intValue")
            | ("Long", "intValue")
            | ("Long", "longValue")
            | ("Double", "doubleValue")
            | ("Double", "intValue")
            | ("Collection", "stream")
            | ("java.util.Collection", "stream")
    ) {
        let direct_callee = match (obj_name.as_str(), method.as_str()) {
            ("String", "length") => Some("__j_str_length"),
            ("String", "trim") | ("String", "strip") => Some("__j_str_trim"),
            ("String", "toUpperCase") => Some("__j_str_to_upper"),
            ("String", "toLowerCase") => Some("__j_str_to_lower"),
            _ => None,
        };
        if let Some(callee) = direct_callee {
            return Ok(Expression::new(ExprKind::Lambda {
                params: vec![Param {
                    name: "__value__".to_string(),
                    type_hint: None,
                    default: None,
                    pass_by: PassBy::Value,
                    is_rest: false,
                    is_kwargs: false,
                    is_optional: false,
                    is_nullable: false,
                }],
                body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident(callee)),
                    args: vec![Argument::positional(Expression::ident("__value__"))],
                    optional: false,
                }))),
                is_async: false,
                captures: vec![],
            }));
        }
        return Ok(Expression::new(ExprKind::Lambda {
            params: vec![Param {
                name: "__value__".to_string(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            }],
            body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(Expression::ident("__value__")),
                    field: method,
                    null_safe: false,
                })),
                args: vec![],
                optional: false,
            }))),
            is_async: false,
            captures: vec![],
        }));
    }

    if obj_name
        .chars()
        .next()
        .is_some_and(|ch| ch.is_ascii_lowercase() || ch == '_')
    {
        if method == "add" {
            return Ok(Expression::new(ExprKind::Lambda {
                params: vec![Param {
                    name: "__value__".to_string(),
                    type_hint: None,
                    default: None,
                    pass_by: PassBy::Value,
                    is_rest: false,
                    is_kwargs: false,
                    is_optional: false,
                    is_nullable: false,
                }],
                body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Member {
                        object: Box::new(Expression::ident(&obj_name)),
                        field: method,
                        null_safe: false,
                    })),
                    args: vec![Argument::positional(Expression::ident("__value__"))],
                    optional: false,
                }))),
                is_async: false,
                captures: vec![],
            }));
        }

        if matches!(
            method.as_str(),
            "length" | "toUpperCase" | "toLowerCase" | "trim" | "strip" | "getValue"
        ) {
            return Ok(Expression::new(ExprKind::Lambda {
                params: vec![],
                body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Member {
                        object: Box::new(Expression::ident(&obj_name)),
                        field: method,
                        null_safe: false,
                    })),
                    args: vec![],
                    optional: false,
                }))),
                is_async: false,
                captures: vec![],
            }));
        }
    }

    Ok(Expression::new(ExprKind::Member {
        object: Box::new(obj_expr),
        field: method,
        null_safe: false,
    }))
}

fn walk_lambda(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner();
    let params_pair = inner.next().ok_or("lambda: missing params")?;
    let params = walk_lambda_params(params_pair)?;
    let mut body_pair = inner.next().ok_or("lambda: missing body")?;
    if body_pair.as_rule() == Rule::lambda_body {
        body_pair = body_pair
            .into_inner()
            .next()
            .ok_or("lambda: missing lambda body")?;
    }
    let body = match body_pair.as_rule() {
        Rule::function_body_block => LambdaBody::Block(walk_block(body_pair)?),
        _ => LambdaBody::Expr(Box::new(walk_expression(body_pair)?)),
    };
    Ok(Expression::new(ExprKind::Lambda {
        params,
        body,
        is_async: false,
        captures: vec![],
    }))
}

fn walk_lambda_params(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    let mut params = Vec::new();
    match pair.as_rule() {
        Rule::lambda_params => {
            for p in pair.into_inner() {
                match p.as_rule() {
                    Rule::typed_lambda_param_list => {
                        for tp in p.into_inner() {
                            if tp.as_rule() == Rule::typed_lambda_param {
                                let mut ti = tp.into_inner().peekable();
                                if ti.peek().map(|x| x.as_rule()) == Some(Rule::final_kw) {
                                    ti.next();
                                }
                                let type_hint =
                                    if ti.peek().map(|x| x.as_rule()) == Some(Rule::type_ref) {
                                        Some(extract_ref_name(&ti.next().unwrap()))
                                    } else {
                                        None
                                    };
                                let name = ti
                                    .next()
                                    .ok_or("typed lambda param: missing name")?
                                    .as_str()
                                    .to_string();
                                params.push(Param {
                                    name,
                                    type_hint,
                                    default: None,
                                    pass_by: PassBy::Value,
                                    is_rest: false,
                                    is_kwargs: false,
                                    is_optional: false,
                                    is_nullable: false,
                                });
                            }
                        }
                    }
                    Rule::ident_name_list => {
                        for ip in p.into_inner() {
                            if ip.as_rule() == Rule::ident_name {
                                params.push(Param {
                                    name: ip.as_str().to_string(),
                                    type_hint: None,
                                    default: None,
                                    pass_by: PassBy::Value,
                                    is_rest: false,
                                    is_kwargs: false,
                                    is_optional: false,
                                    is_nullable: false,
                                });
                            }
                        }
                    }
                    Rule::ident_name => {
                        params.push(Param {
                            name: p.as_str().to_string(),
                            type_hint: None,
                            default: None,
                            pass_by: PassBy::Value,
                            is_rest: false,
                            is_kwargs: false,
                            is_optional: false,
                            is_nullable: false,
                        });
                    }
                    _ => {}
                }
            }
        }
        Rule::ident_name => {
            params.push(Param {
                name: pair.as_str().to_string(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            });
        }
        _ => {}
    }
    Ok(params)
}

fn walk_arguments(pair: Pair<Rule>) -> Result<Vec<Argument>, String> {
    let mut args = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::argument {
            let mut ai = p.into_inner();
            let first = ai.next().ok_or("arg: empty")?;
            if first.as_rule() == Rule::spread_arg {
                let e = walk_expression(first.into_inner().next().ok_or("spread: empty")?)?;
                args.push(Argument {
                    value: e,
                    name: None,
                    by_ref: false,
                    spread: true,
                });
            } else {
                args.push(Argument::positional(walk_expression(first)?));
            }
        }
    }
    Ok(args)
}

fn walk_literal(pair: Pair<Rule>) -> Result<Expression, String> {
    let inner = pair.into_inner().next().ok_or("literal: empty")?;
    match inner.as_rule() {
        Rule::true_kw => Ok(Expression::bool(true)),
        Rule::false_kw => Ok(Expression::bool(false)),
        Rule::null_kw => Ok(Expression::null()),
        Rule::int_literal => {
            let s = inner.as_str().replace('_', "");
            let v = if s.starts_with("0x") || s.starts_with("0X") {
                i64::from_str_radix(&s[2..], 16).unwrap_or(0)
            } else if s.starts_with("0b") || s.starts_with("0B") {
                i64::from_str_radix(&s[2..], 2).unwrap_or(0)
            } else if s.len() > 1 && s.starts_with('0') {
                i64::from_str_radix(&s[1..], 8).unwrap_or(0)
            } else {
                s.parse::<i64>().unwrap_or(0)
            };
            Ok(Expression::int(v))
        }
        Rule::long_literal => {
            let s = inner.as_str().replace('_', "");
            let s = s.trim_end_matches(|c| c == 'l' || c == 'L');
            let v = if s.starts_with("0x") || s.starts_with("0X") {
                i64::from_str_radix(&s[2..], 16).unwrap_or(0)
            } else if s.starts_with("0b") || s.starts_with("0B") {
                i64::from_str_radix(&s[2..], 2).unwrap_or(0)
            } else if s.len() > 1 && s.starts_with('0') {
                i64::from_str_radix(&s[1..], 8).unwrap_or(0)
            } else {
                s.parse::<i64>().unwrap_or(0)
            };
            Ok(Expression::new(ExprKind::Cast {
                type_name: "long".to_string(),
                expr: Box::new(Expression::int(v)),
            }))
        }
        Rule::float_literal => {
            let s = inner.as_str().replace('_', "");
            let type_name = if s.ends_with('f') || s.ends_with('F') {
                Some("float")
            } else {
                None
            };
            let s = s.trim_end_matches(|c| matches!(c, 'f' | 'F' | 'd' | 'D'));
            let expr = Expression::float(s.parse().unwrap_or(0.0));
            if let Some(type_name) = type_name {
                Ok(Expression::new(ExprKind::Cast {
                    type_name: type_name.to_string(),
                    expr: Box::new(expr),
                }))
            } else {
                Ok(expr)
            }
        }
        Rule::char_literal => {
            let s = inner.as_str();
            let content = &s[1..s.len() - 1];
            if let Some(code) = java_unicode_escape_code_unit(content) {
                if (0xD800..0xE000).contains(&code) {
                    return Ok(Expression::int(code as i64));
                }
            }
            Ok(Expression::string(&unescape_java_string(content)))
        }
        Rule::string_literal => {
            let s = inner.as_str();
            Ok(Expression::string(&unescape_java_string(
                &s[1..s.len() - 1],
            )))
        }
        Rule::text_block => {
            let s = inner.as_str();
            let raw = s.trim_start_matches("\"\"\"").trim_end_matches("\"\"\"");
            Ok(Expression::string(&java_text_block_content(raw)))
        }
        _ => Ok(Expression::null()),
    }
}

// ════════════════════════════════════════════════════════════════════════════
// Normalisations
// ════════════════════════════════════════════════════════════════════════════

/// Inject implicit `super()` into child-class constructors that don't already
/// start with an explicit super/base call.
fn inject_implicit_super(members: &mut Vec<ClassMember>) {
    for m in members.iter_mut() {
        if let ClassMember::Constructor {
            base_args, body, ..
        } = m
        {
            if base_args.is_none() {
                let already_has_super = body
                    .first()
                    .map(|s| match &s.kind {
                        StmtKind::Expr(e) => {
                            matches!(
                                &e.kind,
                                ExprKind::Call { callee, .. }
                                    if matches!(&callee.kind, ExprKind::Super)
                            ) || matches!(&e.kind, ExprKind::SuperCall { .. })
                        }
                        _ => false,
                    })
                    .unwrap_or(false);
                if !already_has_super {
                    *base_args = Some(vec![]);
                }
            }
        }
    }
}

fn inject_java_thread_stamps(members: &mut Vec<ClassMember>) {
    rewrite_java_thread_member_bare_calls(members);
    let has_ctor = members
        .iter()
        .any(|member| matches!(member, ClassMember::Constructor { .. }));
    if !has_ctor {
        members.insert(
            0,
            ClassMember::Constructor {
                name: None,
                params: Vec::new(),
                body: vec![java_thread_init_stmt(None)],
                base_args: None,
                initializer_target: ConstructorInitializerTarget::Base,
                visibility: Visibility::Public,
            },
        );
        return;
    }
    for member in members {
        if let ClassMember::Constructor {
            body,
            base_args,
            initializer_target,
            ..
        } = member
        {
            let name_arg = if *initializer_target == ConstructorInitializerTarget::Base {
                base_args.as_ref().and_then(|args| args.first().cloned())
            } else {
                None
            };
            if *initializer_target == ConstructorInitializerTarget::Base {
                *base_args = None;
            }
            body.insert(0, java_thread_init_stmt(name_arg));
        }
    }
}

fn rewrite_java_thread_member_bare_calls(members: &mut [ClassMember]) {
    for member in members {
        match member {
            ClassMember::Method(method) => rewrite_java_thread_bare_calls_stmt(method),
            ClassMember::Constructor { body, .. } => {
                for stmt in body {
                    rewrite_java_thread_bare_calls_stmt(stmt);
                }
            }
            _ => {}
        }
    }
}

fn rewrite_java_thread_bare_calls_stmt(stmt: &mut Statement) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            rewrite_java_thread_bare_calls_expr(expr);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_java_thread_bare_calls_expr(init);
                }
            }
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                rewrite_java_thread_bare_calls_expr(target);
            }
            rewrite_java_thread_bare_calls_expr(value);
        }
        StmtKind::Block(body) | StmtKind::NamespaceDecl { body, .. } => {
            for nested in body {
                rewrite_java_thread_bare_calls_stmt(nested);
            }
        }
        StmtKind::FunctionDecl { body, .. } => {
            for nested in body {
                rewrite_java_thread_bare_calls_stmt(nested);
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_java_thread_bare_calls_expr(cond);
            for nested in then_body {
                rewrite_java_thread_bare_calls_stmt(nested);
            }
            for (elif_cond, elif_body) in elifs {
                rewrite_java_thread_bare_calls_expr(elif_cond);
                for nested in elif_body {
                    rewrite_java_thread_bare_calls_stmt(nested);
                }
            }
            if let Some(else_body) = else_body {
                for nested in else_body {
                    rewrite_java_thread_bare_calls_stmt(nested);
                }
            }
        }
        StmtKind::While { cond, body, .. } | StmtKind::DoWhile { cond, body, .. } => {
            rewrite_java_thread_bare_calls_expr(cond);
            for nested in body {
                rewrite_java_thread_bare_calls_stmt(nested);
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                rewrite_java_thread_bare_calls_stmt(init);
            }
            if let Some(cond) = cond {
                rewrite_java_thread_bare_calls_expr(cond);
            }
            if let Some(update) = update {
                rewrite_java_thread_bare_calls_expr(update);
            }
            for nested in body {
                rewrite_java_thread_bare_calls_stmt(nested);
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            for nested in body {
                rewrite_java_thread_bare_calls_stmt(nested);
            }
            for catch in catches {
                for nested in &mut catch.body {
                    rewrite_java_thread_bare_calls_stmt(nested);
                }
            }
            if let Some(else_body) = else_body {
                for nested in else_body {
                    rewrite_java_thread_bare_calls_stmt(nested);
                }
            }
            if let Some(finally) = finally {
                for nested in finally {
                    rewrite_java_thread_bare_calls_stmt(nested);
                }
            }
        }
        _ => {}
    }
}

fn rewrite_java_thread_bare_calls_expr(expr: &mut Expression) {
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            rewrite_java_thread_bare_calls_expr(callee);
            for arg in &mut *args {
                rewrite_java_thread_bare_calls_expr(&mut arg.value);
            }
            if let ExprKind::Ident(name) = &callee.kind {
                let prelude_fn = match name.as_str() {
                    "start" => Some("__j_thread_start"),
                    "join" => Some("__j_thread_join"),
                    "isAlive" => Some("__j_thread_is_alive"),
                    "getName" => Some("__j_thread_get_name"),
                    "setName" => Some("__j_thread_set_name"),
                    "getPriority" => Some("__j_thread_get_priority"),
                    "setPriority" => Some("__j_thread_set_priority"),
                    "interrupt" => Some("__j_thread_interrupt"),
                    "isInterrupted" => Some("__j_thread_is_interrupted"),
                    _ => None,
                };
                if let Some(prelude_fn) = prelude_fn {
                    *callee = Box::new(Expression::ident(prelude_fn));
                    args.insert(0, Argument::positional(Expression::new(ExprKind::This)));
                }
            }
        }
        ExprKind::Binary { left, right, .. } => {
            rewrite_java_thread_bare_calls_expr(left);
            rewrite_java_thread_bare_calls_expr(right);
        }
        ExprKind::Member { object, .. } => rewrite_java_thread_bare_calls_expr(object),
        ExprKind::Index { object, index, .. } => {
            rewrite_java_thread_bare_calls_expr(object);
            rewrite_java_thread_bare_calls_expr(index);
        }
        ExprKind::Array(elems) => {
            for elem in elems {
                rewrite_java_thread_bare_calls_expr(&mut elem.value);
            }
        }
        ExprKind::Lambda { .. } => {}
        _ => {}
    }
}

fn java_thread_init_stmt(name_arg: Option<Expression>) -> Statement {
    Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("__j_thread_init")),
        args: vec![
            Argument::positional(Expression::new(ExprKind::This)),
            Argument::positional(Expression::null()),
            Argument::positional(
                name_arg.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Undefined))),
            ),
        ],
        optional: false,
    })))
}

/// Extract an explicit `super(...)` / `this(...)` call from the top of a
/// constructor body and put the args in `base_args`.
fn extract_base_call_from_body(
    body: &mut Vec<Statement>,
    base_args: &mut Option<Vec<Expression>>,
    initializer_target: &mut ConstructorInitializerTarget,
) {
    if body.is_empty() {
        return;
    }
    let target = match &body[0].kind {
        StmtKind::Expr(e) => match &e.kind {
            ExprKind::SuperCall { .. } => Some(ConstructorInitializerTarget::Base),
            ExprKind::Call { callee, .. } => match &callee.kind {
                ExprKind::Super => Some(ConstructorInitializerTarget::Base),
                ExprKind::This => Some(ConstructorInitializerTarget::This),
                _ => None,
            },
            _ => None,
        },
        _ => None,
    };
    if let Some(target) = target {
        let s = body.remove(0);
        if let StmtKind::Expr(e) = s.kind {
            let args_exprs: Vec<Expression> = match e.kind {
                ExprKind::SuperCall { args, .. } => args.into_iter().map(|a| a.value).collect(),
                ExprKind::Call { args, .. } => args.into_iter().map(|a| a.value).collect(),
                _ => vec![],
            };
            *base_args = Some(args_exprs);
            *initializer_target = target;
        }
    }
}

fn default_expr_for_java_type(type_name: &str) -> Option<Expression> {
    match type_name {
        "byte" | "short" | "int" | "long" | "char" => Some(Expression::int(0)),
        "float" | "double" => Some(Expression::float(0.0)),
        "boolean" => Some(Expression::bool(false)),
        _ => None,
    }
}

// ════════════════════════════════════════════════════════════════════════════
// Local (method-body) classes → static nested siblings
// ════════════════════════════════════════════════════════════════════════════

fn hoist_java_local_classes(body: &mut [Statement]) {
    for stmt in body {
        hoist_java_local_classes_stmt(stmt);
    }
}

fn hoist_java_local_classes_stmt(stmt: &mut Statement) {
    match &mut stmt.kind {
        StmtKind::ClassDecl { name, members, .. } => {
            let class_name = name.clone();
            let mut hoisted: Vec<ClassMember> = Vec::new();
            let mut counter = 0usize;
            for member in members.iter_mut() {
                match member {
                    ClassMember::Method(m) => {
                        if let StmtKind::FunctionDecl { params, body, .. } = &mut m.kind {
                            let mut scope: HashSet<String> =
                                params.iter().map(|p| p.name.clone()).collect();
                            hoist_local_classes_in_body(
                                body,
                                &mut scope,
                                &class_name,
                                &mut counter,
                                &mut hoisted,
                            );
                        }
                    }
                    ClassMember::Constructor { params, body, .. } => {
                        let mut scope: HashSet<String> =
                            params.iter().map(|p| p.name.clone()).collect();
                        hoist_local_classes_in_body(
                            body,
                            &mut scope,
                            &class_name,
                            &mut counter,
                            &mut hoisted,
                        );
                    }
                    ClassMember::NestedType(ns) => hoist_java_local_classes_stmt(ns),
                    _ => {}
                }
            }
            members.extend(hoisted);
        }
        StmtKind::Block(b) | StmtKind::NamespaceDecl { body: b, .. } => hoist_java_local_classes(b),
        _ => {}
    }
}

/// Process one statement list: hoist any local class declarations out to
/// `hoisted` (as static nested siblings), thread captured enclosing locals
/// through their constructors, and rewrite `new Local(...)` sites in this
/// body to the hoisted name with capture values appended.
fn hoist_local_classes_in_body(
    body: &mut Vec<Statement>,
    scope: &mut HashSet<String>,
    enclosing: &str,
    counter: &mut usize,
    hoisted: &mut Vec<ClassMember>,
) {
    // (old local name, hoisted name, capture value exprs) for `new` rewrite.
    let mut mappings: Vec<(String, String, Vec<Expression>)> = Vec::new();

    for stmt in body.iter_mut() {
        match &mut stmt.kind {
            StmtKind::VarDecl { declarations, .. } => {
                for decl in declarations {
                    collect_binding_names(&decl.pattern, scope);
                }
            }
            StmtKind::ClassDecl {
                name: local_name,
                members: local_members,
                parents,
                ..
            } => {
                // Captured enclosing locals: free idents used in the class
                // body that are in scope and not the class's own members.
                let mut used: HashSet<String> = HashSet::new();
                java_collect_member_idents(local_members, &mut used);
                let own: HashSet<String> = local_members
                    .iter()
                    .filter_map(|m| match m {
                        ClassMember::Field { name, .. } => Some(name.clone()),
                        _ => None,
                    })
                    .collect();
                let mut captures: Vec<String> = used
                    .into_iter()
                    .filter(|n| scope.contains(n) && !own.contains(n))
                    .collect();
                captures.sort();

                let new_name = format!("{enclosing}__local{counter}__{local_name}");
                *counter += 1;
                let old_name = local_name.clone();

                if !captures.is_empty() {
                    java_thread_local_class_captures(local_members, &captures);
                }
                let capture_vals: Vec<Expression> =
                    captures.iter().map(|c| Expression::ident(c)).collect();

                // Build the hoisted static nested class.
                let mut modifiers = ClassModifiers::default();
                modifiers.is_static = true;
                let hoisted_stmt = Statement::new(StmtKind::ClassDecl {
                    name: new_name.clone(),
                    parents: std::mem::take(parents),
                    interfaces: vec![],
                    members: std::mem::take(local_members),
                    modifiers,
                    decorators: vec![],
                });
                hoisted.push(ClassMember::NestedType(Box::new(hoisted_stmt)));
                mappings.push((old_name, new_name, capture_vals));

                // Blank the original declaration statement.
                stmt.kind = StmtKind::Block(vec![]);
            }
            // Recurse into nested control-flow bodies with the current scope.
            StmtKind::Block(b) | StmtKind::NamespaceDecl { body: b, .. } => {
                hoist_local_classes_in_body(b, &mut scope.clone(), enclosing, counter, hoisted);
            }
            StmtKind::For { body: b, .. }
            | StmtKind::While { body: b, .. }
            | StmtKind::DoWhile { body: b, .. }
            | StmtKind::ForIn { body: b, .. } => {
                hoist_local_classes_in_body(b, &mut scope.clone(), enclosing, counter, hoisted);
            }
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                hoist_local_classes_in_body(
                    then_body,
                    &mut scope.clone(),
                    enclosing,
                    counter,
                    hoisted,
                );
                for (_, eb) in elifs {
                    hoist_local_classes_in_body(
                        eb,
                        &mut scope.clone(),
                        enclosing,
                        counter,
                        hoisted,
                    );
                }
                if let Some(eb) = else_body {
                    hoist_local_classes_in_body(
                        eb,
                        &mut scope.clone(),
                        enclosing,
                        counter,
                        hoisted,
                    );
                }
            }
            _ => {}
        }
    }

    // Rewrite construction sites for each hoisted local class.
    for (old_name, new_name, capture_vals) in &mappings {
        for stmt in body.iter_mut() {
            rewrite_java_new_local_stmt(stmt, old_name, new_name, capture_vals);
        }
    }
}

/// Add capture fields + constructor threading to a local class for each
/// captured enclosing local. A field with the captured name is added (so
/// existing bare-name → `this.field` qualification handles method refs), and
/// every constructor (or a synthesized one) takes the value as a trailing
/// parameter and stores it.
fn java_thread_local_class_captures(members: &mut Vec<ClassMember>, captures: &[String]) {
    for name in captures {
        members.insert(
            0,
            ClassMember::Field {
                name: name.clone(),
                type_hint: None,
                init: None,
                modifiers: Modifiers::default(),
                with_events: false,
                array_bounds: None,
            },
        );
    }
    // The ctor param must NOT share the capture field's name — the
    // bare-field → `this.field` qualification pass would otherwise rewrite
    // the store's RHS to `this.<name>` (self-assigning the uninitialized
    // field). Java relies on param-shadowing here; we use a distinct name.
    let capture_param = |name: &str| Param {
        name: format!("__cap_{name}"),
        type_hint: None,
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    };
    let store = |name: &str| {
        Statement::new(StmtKind::Assign {
            targets: vec![Expression::new(ExprKind::Member {
                object: Box::new(Expression::new(ExprKind::This)),
                field: name.to_string(),
                null_safe: false,
            })],
            value: Expression::ident(&format!("__cap_{name}")),
        })
    };
    let mut saw_ctor = false;
    for member in members.iter_mut() {
        if let ClassMember::Constructor { params, body, .. } = member {
            saw_ctor = true;
            for name in captures {
                params.push(capture_param(name));
            }
            let mut prelude: Vec<Statement> = captures.iter().map(|n| store(n)).collect();
            prelude.append(body);
            *body = prelude;
        }
    }
    if !saw_ctor {
        members.push(ClassMember::Constructor {
            name: None,
            params: captures.iter().map(|n| capture_param(n)).collect(),
            body: captures.iter().map(|n| store(n)).collect(),
            base_args: None,
            initializer_target: ConstructorInitializerTarget::Base,
            visibility: ParsedModifiers::default().visibility,
        });
    }
}

/// Collect every identifier referenced anywhere in a class's members.
fn java_collect_member_idents(members: &[ClassMember], out: &mut HashSet<String>) {
    for member in members {
        match member {
            ClassMember::Field { init: Some(e), .. } => java_collect_idents_expr(e, out),
            ClassMember::Method(m) => {
                if let StmtKind::FunctionDecl { body, .. } = &m.kind {
                    java_collect_idents_stmts(body, out);
                }
            }
            ClassMember::Constructor { body, .. } => java_collect_idents_stmts(body, out),
            _ => {}
        }
    }
}

fn java_collect_idents_stmts(stmts: &[Statement], out: &mut HashSet<String>) {
    for stmt in stmts {
        java_collect_idents_stmt(stmt, out);
    }
}

fn java_collect_idents_stmt(stmt: &Statement, out: &mut HashSet<String>) {
    match &stmt.kind {
        StmtKind::Expr(e) | StmtKind::Return(Some(e)) | StmtKind::Throw { expr: Some(e), .. } => {
            java_collect_idents_expr(e, out)
        }
        StmtKind::VarDecl { declarations, .. } => {
            for d in declarations {
                if let Some(e) = &d.init {
                    java_collect_idents_expr(e, out);
                }
            }
        }
        StmtKind::Assign { targets, value } => {
            for t in targets {
                java_collect_idents_expr(t, out);
            }
            java_collect_idents_expr(value, out);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            java_collect_idents_expr(target, out);
            java_collect_idents_expr(value, out);
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            java_collect_idents_expr(cond, out);
            java_collect_idents_stmts(then_body, out);
            for (c, b) in elifs {
                java_collect_idents_expr(c, out);
                java_collect_idents_stmts(b, out);
            }
            if let Some(b) = else_body {
                java_collect_idents_stmts(b, out);
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(i) = init {
                java_collect_idents_stmt(i, out);
            }
            if let Some(c) = cond {
                java_collect_idents_expr(c, out);
            }
            if let Some(u) = update {
                java_collect_idents_expr(u, out);
            }
            java_collect_idents_stmts(body, out);
        }
        StmtKind::While { cond, body, .. } | StmtKind::DoWhile { cond, body, .. } => {
            java_collect_idents_expr(cond, out);
            java_collect_idents_stmts(body, out);
        }
        StmtKind::ForIn { iter, body, .. } => {
            java_collect_idents_expr(iter, out);
            java_collect_idents_stmts(body, out);
        }
        StmtKind::Block(b) => java_collect_idents_stmts(b, out),
        _ => {}
    }
}

fn java_collect_idents_expr(expr: &Expression, out: &mut HashSet<String>) {
    match &expr.kind {
        ExprKind::Ident(name) => {
            out.insert(name.clone());
        }
        ExprKind::Binary { left, right, .. } => {
            java_collect_idents_expr(left, out);
            java_collect_idents_expr(right, out);
        }
        ExprKind::Unary { expr: e, .. }
        | ExprKind::Spread(e)
        | ExprKind::Await(e)
        | ExprKind::TypeOf(e) => java_collect_idents_expr(e, out),
        ExprKind::Ternary { cond, then, else_ } => {
            java_collect_idents_expr(cond, out);
            java_collect_idents_expr(then, out);
            java_collect_idents_expr(else_, out);
        }
        ExprKind::Member { object, .. } => java_collect_idents_expr(object, out),
        ExprKind::Index { object, index, .. } => {
            java_collect_idents_expr(object, out);
            java_collect_idents_expr(index, out);
        }
        ExprKind::Call { callee, args, .. } => {
            java_collect_idents_expr(callee, out);
            for a in args {
                java_collect_idents_expr(&a.value, out);
            }
        }
        ExprKind::New { class, args } => {
            java_collect_idents_expr(class, out);
            for a in args {
                java_collect_idents_expr(&a.value, out);
            }
        }
        ExprKind::Assign { target, value } => {
            java_collect_idents_expr(target, out);
            java_collect_idents_expr(value, out);
        }
        ExprKind::Array(elems) => {
            for e in elems {
                java_collect_idents_expr(&e.value, out);
            }
        }
        _ => {}
    }
}

/// Rewrite `new OldLocal(args)` → `new HoistedName(args, capture_vals…)`
/// everywhere in a statement.
fn rewrite_java_new_local_stmt(
    stmt: &mut Statement,
    old_name: &str,
    new_name: &str,
    capture_vals: &[Expression],
) {
    match &mut stmt.kind {
        StmtKind::Expr(e) | StmtKind::Return(Some(e)) | StmtKind::Throw { expr: Some(e), .. } => {
            rewrite_java_new_local_expr(e, old_name, new_name, capture_vals)
        }
        StmtKind::VarDecl { declarations, .. } => {
            for d in declarations {
                // Retype `Local loc = …` → the hoisted name so `loc.method()`
                // resolves to the user class (not a builtin value-method).
                if d.type_hint.as_deref() == Some(old_name) {
                    d.type_hint = Some(new_name.to_string());
                }
                if let Some(e) = &mut d.init {
                    rewrite_java_new_local_expr(e, old_name, new_name, capture_vals);
                }
            }
        }
        StmtKind::Assign { targets, value } => {
            for t in targets {
                rewrite_java_new_local_expr(t, old_name, new_name, capture_vals);
            }
            rewrite_java_new_local_expr(value, old_name, new_name, capture_vals);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_java_new_local_expr(target, old_name, new_name, capture_vals);
            rewrite_java_new_local_expr(value, old_name, new_name, capture_vals);
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_java_new_local_expr(cond, old_name, new_name, capture_vals);
            for s in then_body.iter_mut() {
                rewrite_java_new_local_stmt(s, old_name, new_name, capture_vals);
            }
            for (c, b) in elifs {
                rewrite_java_new_local_expr(c, old_name, new_name, capture_vals);
                for s in b.iter_mut() {
                    rewrite_java_new_local_stmt(s, old_name, new_name, capture_vals);
                }
            }
            if let Some(b) = else_body {
                for s in b.iter_mut() {
                    rewrite_java_new_local_stmt(s, old_name, new_name, capture_vals);
                }
            }
        }
        StmtKind::For {
            cond, update, body, ..
        } => {
            if let Some(c) = cond {
                rewrite_java_new_local_expr(c, old_name, new_name, capture_vals);
            }
            if let Some(u) = update {
                rewrite_java_new_local_expr(u, old_name, new_name, capture_vals);
            }
            for s in body.iter_mut() {
                rewrite_java_new_local_stmt(s, old_name, new_name, capture_vals);
            }
        }
        StmtKind::While { cond, body, .. } | StmtKind::DoWhile { cond, body, .. } => {
            rewrite_java_new_local_expr(cond, old_name, new_name, capture_vals);
            for s in body.iter_mut() {
                rewrite_java_new_local_stmt(s, old_name, new_name, capture_vals);
            }
        }
        StmtKind::ForIn { iter, body, .. } => {
            rewrite_java_new_local_expr(iter, old_name, new_name, capture_vals);
            for s in body.iter_mut() {
                rewrite_java_new_local_stmt(s, old_name, new_name, capture_vals);
            }
        }
        StmtKind::Block(b) => {
            for s in b.iter_mut() {
                rewrite_java_new_local_stmt(s, old_name, new_name, capture_vals);
            }
        }
        _ => {}
    }
}

fn rewrite_java_new_local_expr(
    expr: &mut Expression,
    old_name: &str,
    new_name: &str,
    capture_vals: &[Expression],
) {
    match &mut expr.kind {
        ExprKind::New { class, args } => {
            if let ExprKind::Ident(n) = &class.kind {
                if n == old_name {
                    class.kind = ExprKind::Ident(new_name.to_string());
                    for v in capture_vals {
                        args.push(Argument::positional(v.clone()));
                    }
                }
            }
            for a in args.iter_mut() {
                rewrite_java_new_local_expr(&mut a.value, old_name, new_name, capture_vals);
            }
        }
        ExprKind::Binary { left, right, .. } => {
            rewrite_java_new_local_expr(left, old_name, new_name, capture_vals);
            rewrite_java_new_local_expr(right, old_name, new_name, capture_vals);
        }
        ExprKind::Unary { expr: e, .. }
        | ExprKind::Spread(e)
        | ExprKind::Await(e)
        | ExprKind::TypeOf(e) => rewrite_java_new_local_expr(e, old_name, new_name, capture_vals),
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_java_new_local_expr(cond, old_name, new_name, capture_vals);
            rewrite_java_new_local_expr(then, old_name, new_name, capture_vals);
            rewrite_java_new_local_expr(else_, old_name, new_name, capture_vals);
        }
        ExprKind::Member { object, .. } => {
            rewrite_java_new_local_expr(object, old_name, new_name, capture_vals)
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_java_new_local_expr(object, old_name, new_name, capture_vals);
            rewrite_java_new_local_expr(index, old_name, new_name, capture_vals);
        }
        ExprKind::Call { callee, args, .. } => {
            rewrite_java_new_local_expr(callee, old_name, new_name, capture_vals);
            for a in args.iter_mut() {
                rewrite_java_new_local_expr(&mut a.value, old_name, new_name, capture_vals);
            }
        }
        ExprKind::Assign { target, value } => {
            rewrite_java_new_local_expr(target, old_name, new_name, capture_vals);
            rewrite_java_new_local_expr(value, old_name, new_name, capture_vals);
        }
        ExprKind::Array(elems) => {
            for e in elems.iter_mut() {
                rewrite_java_new_local_expr(&mut e.value, old_name, new_name, capture_vals);
            }
        }
        _ => {}
    }
}

fn qualify_java_nested_types(body: &mut [Statement]) {
    for stmt in body {
        qualify_java_nested_types_stmt(stmt, None);
    }
}

fn qualify_java_nested_types_stmt(stmt: &mut Statement, owner: Option<&str>) {
    match &mut stmt.kind {
        StmtKind::ClassDecl { name, members, .. } | StmtKind::StructDecl { name, members, .. } => {
            if let Some(owner_name) = owner.filter(|owner_name| *owner_name != "Main") {
                if !name.contains('.') {
                    *name = format!("{}.{}", owner_name, name);
                }
            }
            if name.contains('.') {
                JAVA_NESTED_TYPE_NAMES.with(|names| {
                    names.borrow_mut().insert(name.clone());
                });
            }
            let current_name = name.clone();
            let owner_fields = java_instance_field_names(members);
            let owner_static_fields = java_static_field_names(members);
            let owner_static_field_inits = java_static_field_inits(members);
            let owner_methods = java_instance_method_names(members);
            for member in &mut *members {
                if let ClassMember::NestedType(nested) = member {
                    qualify_java_nested_types_stmt(nested, Some(&current_name));
                }
            }
            let nested_types = java_immediate_nested_types(members);
            for member in &mut *members {
                if let ClassMember::NestedType(nested) = member {
                    rewrite_java_outer_static_refs_nested(
                        nested,
                        &current_name,
                        &owner_static_fields,
                        &owner_static_field_inits,
                    );
                    adapt_java_inner_class(nested, &current_name, &owner_fields, &owner_methods);
                }
            }
            rewrite_java_nested_type_refs_in_members(members, &nested_types);
        }
        StmtKind::InterfaceDecl { name, .. } => {
            if let Some(owner_name) = owner.filter(|owner_name| *owner_name != "Main") {
                if !name.contains('.') {
                    *name = format!("{}.{}", owner_name, name);
                }
            }
            if name.contains('.') {
                JAVA_NESTED_TYPE_NAMES.with(|names| {
                    names.borrow_mut().insert(name.clone());
                });
            }
        }
        StmtKind::EnumDecl {
            name, body_members, ..
        } => {
            if let Some(owner_name) = owner.filter(|owner_name| *owner_name != "Main") {
                if !name.contains('.') {
                    *name = format!("{}.{}", owner_name, name);
                }
            }
            if name.contains('.') {
                JAVA_NESTED_TYPE_NAMES.with(|names| {
                    names.borrow_mut().insert(name.clone());
                });
            }
            let current_name = name.clone();
            for member in body_members {
                if let ClassMember::NestedType(nested) = member {
                    qualify_java_nested_types_stmt(nested, Some(&current_name));
                }
            }
        }
        StmtKind::Block(stmts) | StmtKind::NamespaceDecl { body: stmts, .. } => {
            for stmt in stmts {
                qualify_java_nested_types_stmt(stmt, owner);
            }
        }
        _ => {}
    }
}

fn java_instance_field_names(members: &[ClassMember]) -> HashSet<String> {
    members
        .iter()
        .filter_map(|member| match member {
            ClassMember::Field {
                name, modifiers, ..
            } if !modifiers.is_static => Some(name.clone()),
            _ => None,
        })
        .collect()
}

fn java_static_field_names(members: &[ClassMember]) -> HashSet<String> {
    members
        .iter()
        .filter_map(|member| match member {
            ClassMember::Field {
                name, modifiers, ..
            } if modifiers.is_static => Some(name.clone()),
            _ => None,
        })
        .collect()
}

fn java_static_field_inits(members: &[ClassMember]) -> HashMap<String, Expression> {
    members
        .iter()
        .filter_map(|member| match member {
            ClassMember::Field {
                name,
                init: Some(init),
                modifiers,
                ..
            } if modifiers.is_static && java_static_field_init_is_inlineable(init) => {
                Some((name.clone(), init.clone()))
            }
            _ => None,
        })
        .collect()
}

fn java_static_field_init_is_inlineable(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Lit(_) => true,
        ExprKind::Unary { expr, .. } | ExprKind::Cast { expr, .. } => {
            java_static_field_init_is_inlineable(expr)
        }
        ExprKind::Binary { left, right, .. } => {
            java_static_field_init_is_inlineable(left)
                && java_static_field_init_is_inlineable(right)
        }
        ExprKind::Array(elems) => elems
            .iter()
            .all(|elem| java_static_field_init_is_inlineable(&elem.value)),
        _ => false,
    }
}

fn java_instance_method_names(members: &[ClassMember]) -> HashSet<String> {
    members
        .iter()
        .filter_map(|member| match member {
            ClassMember::Method(stmt) => match &stmt.kind {
                StmtKind::FunctionDecl {
                    name, modifiers, ..
                } if !modifiers.is_static => Some(name.clone()),
                _ => None,
            },
            _ => None,
        })
        .collect()
}

fn java_immediate_nested_types(members: &[ClassMember]) -> HashMap<String, (String, bool)> {
    let mut nested = HashMap::new();
    for member in members {
        let ClassMember::NestedType(stmt) = member else {
            continue;
        };
        match &stmt.kind {
            StmtKind::ClassDecl {
                name, modifiers, ..
            } => {
                let simple = name.rsplit('.').next().unwrap_or(name).to_string();
                nested.insert(simple, (name.clone(), modifiers.is_static));
            }
            StmtKind::StructDecl { name, .. }
            | StmtKind::InterfaceDecl { name, .. }
            | StmtKind::EnumDecl { name, .. } => {
                let simple = name.rsplit('.').next().unwrap_or(name).to_string();
                nested.insert(simple, (name.clone(), true));
            }
            _ => {}
        }
    }
    nested
}

fn adapt_java_inner_class(
    nested: &mut Statement,
    owner_name: &str,
    owner_fields: &HashSet<String>,
    owner_methods: &HashSet<String>,
) {
    let StmtKind::ClassDecl {
        members, modifiers, ..
    } = &mut nested.kind
    else {
        return;
    };
    if modifiers.is_static {
        return;
    }

    let inner_fields = java_instance_field_names(members);
    let visible_owner_fields: HashSet<String> =
        owner_fields.difference(&inner_fields).cloned().collect();

    let abstract_without_outer_refs = modifiers.is_abstract
        && !java_inner_class_needs_outer(members, owner_name, &visible_owner_fields, owner_methods);
    if abstract_without_outer_refs {
        return;
    }

    if !members.iter().any(|member| {
        matches!(
            member,
            ClassMember::Field { name, .. } if name == "__java_outer"
        )
    }) {
        members.insert(
            0,
            ClassMember::Field {
                name: "__java_outer".to_string(),
                type_hint: Some(owner_name.to_string()),
                init: None,
                modifiers: Modifiers::default(),
                with_events: false,
                array_bounds: None,
            },
        );
    }

    let mut outer_field_inits = Vec::new();
    for member in members.iter_mut() {
        let ClassMember::Field {
            name,
            init,
            modifiers,
            ..
        } = member
        else {
            continue;
        };
        if modifiers.is_static {
            continue;
        }
        let Some(init_expr) = init else {
            continue;
        };
        rewrite_java_inner_outer_refs_expr(
            init_expr,
            owner_name,
            &visible_owner_fields,
            owner_methods,
        );
        if java_expr_reads_java_outer(init_expr) {
            let value = init_expr.clone();
            *init = None;
            outer_field_inits.push(java_inner_field_init_assign_stmt(name, value));
        }
    }

    let mut saw_constructor = false;
    for member in members.iter_mut() {
        match member {
            ClassMember::Constructor { params, body, .. } => {
                saw_constructor = true;
                java_prepend_outer_constructor_param(params, body, owner_name);
                java_insert_inner_field_inits(body, &outer_field_inits);
                rewrite_java_inner_outer_refs_stmts(
                    body,
                    owner_name,
                    &visible_owner_fields,
                    owner_methods,
                );
            }
            ClassMember::Method(stmt) => {
                rewrite_java_inner_outer_refs_stmt(
                    stmt,
                    owner_name,
                    &visible_owner_fields,
                    owner_methods,
                );
            }
            _ => {}
        }
    }

    if !saw_constructor {
        let mut body = vec![java_outer_assign_stmt()];
        body.extend(outer_field_inits);
        members.push(ClassMember::Constructor {
            name: None,
            params: vec![java_outer_param(owner_name)],
            body,
            base_args: None,
            initializer_target: ConstructorInitializerTarget::Base,
            visibility: Visibility::Public,
        });
    }
}

fn java_inner_class_needs_outer(
    members: &[ClassMember],
    owner_name: &str,
    owner_fields: &HashSet<String>,
    owner_methods: &HashSet<String>,
) -> bool {
    members.iter().any(|member| match member {
        ClassMember::Field {
            init: Some(init),
            modifiers,
            ..
        } if !modifiers.is_static => {
            java_expr_needs_java_outer(init, owner_name, owner_fields, owner_methods)
        }
        ClassMember::Constructor { body, .. } => body
            .iter()
            .any(|stmt| java_stmt_needs_java_outer(stmt, owner_name, owner_fields, owner_methods)),
        ClassMember::Method(stmt) => {
            java_stmt_needs_java_outer(stmt, owner_name, owner_fields, owner_methods)
        }
        _ => false,
    })
}

fn java_stmt_needs_java_outer(
    stmt: &Statement,
    owner_name: &str,
    owner_fields: &HashSet<String>,
    owner_methods: &HashSet<String>,
) -> bool {
    match &stmt.kind {
        StmtKind::FunctionDecl { body, .. } | StmtKind::Block(body) => body
            .iter()
            .any(|stmt| java_stmt_needs_java_outer(stmt, owner_name, owner_fields, owner_methods)),
        StmtKind::VarDecl { declarations, .. } => declarations.iter().any(|decl| {
            decl.init.as_ref().is_some_and(|init| {
                java_expr_needs_java_outer(init, owner_name, owner_fields, owner_methods)
            })
        }),
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            java_expr_needs_java_outer(expr, owner_name, owner_fields, owner_methods)
        }
        StmtKind::Assign { targets, value } => {
            java_expr_needs_java_outer(value, owner_name, owner_fields, owner_methods)
                || targets.iter().any(|target| {
                    java_expr_needs_java_outer(target, owner_name, owner_fields, owner_methods)
                })
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            java_expr_needs_java_outer(target, owner_name, owner_fields, owner_methods)
                || java_expr_needs_java_outer(value, owner_name, owner_fields, owner_methods)
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            java_expr_needs_java_outer(cond, owner_name, owner_fields, owner_methods)
                || then_body.iter().any(|stmt| {
                    java_stmt_needs_java_outer(stmt, owner_name, owner_fields, owner_methods)
                })
                || elifs.iter().any(|(cond, body)| {
                    java_expr_needs_java_outer(cond, owner_name, owner_fields, owner_methods)
                        || body.iter().any(|stmt| {
                            java_stmt_needs_java_outer(
                                stmt,
                                owner_name,
                                owner_fields,
                                owner_methods,
                            )
                        })
                })
                || else_body.as_ref().is_some_and(|body| {
                    body.iter().any(|stmt| {
                        java_stmt_needs_java_outer(stmt, owner_name, owner_fields, owner_methods)
                    })
                })
        }
        StmtKind::While { cond, body, .. } | StmtKind::DoWhile { cond, body, .. } => {
            java_expr_needs_java_outer(cond, owner_name, owner_fields, owner_methods)
                || body.iter().any(|stmt| {
                    java_stmt_needs_java_outer(stmt, owner_name, owner_fields, owner_methods)
                })
        }
        _ => false,
    }
}

fn java_expr_needs_java_outer(
    expr: &Expression,
    owner_name: &str,
    owner_fields: &HashSet<String>,
    owner_methods: &HashSet<String>,
) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => owner_fields.contains(name),
        ExprKind::Call { callee, args, .. } => {
            matches!(&callee.kind, ExprKind::Ident(name) if owner_methods.contains(name))
                || java_expr_needs_java_outer(callee, owner_name, owner_fields, owner_methods)
                || args.iter().any(|arg| {
                    java_expr_needs_java_outer(&arg.value, owner_name, owner_fields, owner_methods)
                })
        }
        ExprKind::Member { object, field, .. } => {
            (field == "this" && java_expr_dotted_name(object).as_deref() == Some(owner_name))
                || java_expr_needs_java_outer(object, owner_name, owner_fields, owner_methods)
        }
        ExprKind::Index { object, index, .. } => {
            java_expr_needs_java_outer(object, owner_name, owner_fields, owner_methods)
                || java_expr_needs_java_outer(index, owner_name, owner_fields, owner_methods)
        }
        ExprKind::Binary { left, right, .. } => {
            java_expr_needs_java_outer(left, owner_name, owner_fields, owner_methods)
                || java_expr_needs_java_outer(right, owner_name, owner_fields, owner_methods)
        }
        ExprKind::Unary { expr, .. } | ExprKind::Cast { expr, .. } => {
            java_expr_needs_java_outer(expr, owner_name, owner_fields, owner_methods)
        }
        ExprKind::Assign { target, value } => {
            java_expr_needs_java_outer(target, owner_name, owner_fields, owner_methods)
                || java_expr_needs_java_outer(value, owner_name, owner_fields, owner_methods)
        }
        ExprKind::Ternary { cond, then, else_ } => {
            java_expr_needs_java_outer(cond, owner_name, owner_fields, owner_methods)
                || java_expr_needs_java_outer(then, owner_name, owner_fields, owner_methods)
                || java_expr_needs_java_outer(else_, owner_name, owner_fields, owner_methods)
        }
        ExprKind::Array(elems) => elems.iter().any(|elem| {
            java_expr_needs_java_outer(&elem.value, owner_name, owner_fields, owner_methods)
        }),
        ExprKind::New { class, args } => {
            java_expr_needs_java_outer(class, owner_name, owner_fields, owner_methods)
                || args.iter().any(|arg| {
                    java_expr_needs_java_outer(&arg.value, owner_name, owner_fields, owner_methods)
                })
        }
        _ => false,
    }
}

fn rewrite_java_outer_static_refs_nested(
    nested: &mut Statement,
    owner_name: &str,
    owner_static_fields: &HashSet<String>,
    owner_static_field_inits: &HashMap<String, Expression>,
) {
    let StmtKind::ClassDecl { members, .. } = &mut nested.kind else {
        return;
    };
    for member in members {
        match member {
            ClassMember::Field {
                init: Some(init), ..
            } => {
                rewrite_java_outer_static_refs_expr(
                    init,
                    owner_name,
                    owner_static_fields,
                    owner_static_field_inits,
                );
            }
            ClassMember::Constructor { body, .. } => {
                rewrite_java_outer_static_refs_stmts(
                    body,
                    owner_name,
                    owner_static_fields,
                    owner_static_field_inits,
                );
            }
            ClassMember::Method(stmt) => {
                rewrite_java_outer_static_refs_stmt(
                    stmt,
                    owner_name,
                    owner_static_fields,
                    owner_static_field_inits,
                );
            }
            ClassMember::NestedType(nested) => {
                rewrite_java_outer_static_refs_nested(
                    nested,
                    owner_name,
                    owner_static_fields,
                    owner_static_field_inits,
                );
            }
            _ => {}
        }
    }
}

fn rewrite_java_outer_static_refs_stmt(
    stmt: &mut Statement,
    owner_name: &str,
    owner_static_fields: &HashSet<String>,
    owner_static_field_inits: &HashMap<String, Expression>,
) {
    match &mut stmt.kind {
        StmtKind::FunctionDecl { body, .. } | StmtKind::Block(body) => {
            rewrite_java_outer_static_refs_stmts(
                body,
                owner_name,
                owner_static_fields,
                owner_static_field_inits,
            );
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_java_outer_static_refs_expr(
                        init,
                        owner_name,
                        owner_static_fields,
                        owner_static_field_inits,
                    );
                }
            }
        }
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            rewrite_java_outer_static_refs_expr(
                expr,
                owner_name,
                owner_static_fields,
                owner_static_field_inits,
            );
        }
        StmtKind::Assign { targets, value } => {
            rewrite_java_outer_static_refs_expr(
                value,
                owner_name,
                owner_static_fields,
                owner_static_field_inits,
            );
            for target in targets {
                rewrite_java_outer_static_refs_expr(
                    target,
                    owner_name,
                    owner_static_fields,
                    owner_static_field_inits,
                );
            }
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_java_outer_static_refs_expr(
                value,
                owner_name,
                owner_static_fields,
                owner_static_field_inits,
            );
            rewrite_java_outer_static_refs_expr(
                target,
                owner_name,
                owner_static_fields,
                owner_static_field_inits,
            );
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_java_outer_static_refs_expr(
                cond,
                owner_name,
                owner_static_fields,
                owner_static_field_inits,
            );
            rewrite_java_outer_static_refs_stmts(
                then_body,
                owner_name,
                owner_static_fields,
                owner_static_field_inits,
            );
            for (elif_cond, elif_body) in elifs {
                rewrite_java_outer_static_refs_expr(
                    elif_cond,
                    owner_name,
                    owner_static_fields,
                    owner_static_field_inits,
                );
                rewrite_java_outer_static_refs_stmts(
                    elif_body,
                    owner_name,
                    owner_static_fields,
                    owner_static_field_inits,
                );
            }
            if let Some(else_body) = else_body {
                rewrite_java_outer_static_refs_stmts(
                    else_body,
                    owner_name,
                    owner_static_fields,
                    owner_static_field_inits,
                );
            }
        }
        StmtKind::While { cond, body, .. } => {
            rewrite_java_outer_static_refs_expr(
                cond,
                owner_name,
                owner_static_fields,
                owner_static_field_inits,
            );
            rewrite_java_outer_static_refs_stmts(
                body,
                owner_name,
                owner_static_fields,
                owner_static_field_inits,
            );
        }
        StmtKind::Try {
            body,
            catches,
            finally,
            ..
        } => {
            rewrite_java_outer_static_refs_stmts(
                body,
                owner_name,
                owner_static_fields,
                owner_static_field_inits,
            );
            for catch in catches {
                rewrite_java_outer_static_refs_stmts(
                    &mut catch.body,
                    owner_name,
                    owner_static_fields,
                    owner_static_field_inits,
                );
            }
            if let Some(finally) = finally {
                rewrite_java_outer_static_refs_stmts(
                    finally,
                    owner_name,
                    owner_static_fields,
                    owner_static_field_inits,
                );
            }
        }
        _ => {}
    }
}

fn rewrite_java_outer_static_refs_stmts(
    stmts: &mut [Statement],
    owner_name: &str,
    owner_static_fields: &HashSet<String>,
    owner_static_field_inits: &HashMap<String, Expression>,
) {
    for stmt in stmts {
        rewrite_java_outer_static_refs_stmt(
            stmt,
            owner_name,
            owner_static_fields,
            owner_static_field_inits,
        );
    }
}

fn rewrite_java_outer_static_refs_expr(
    expr: &mut Expression,
    owner_name: &str,
    owner_static_fields: &HashSet<String>,
    owner_static_field_inits: &HashMap<String, Expression>,
) {
    match &mut expr.kind {
        ExprKind::Ident(name) if owner_static_fields.contains(name) => {
            let field = name.clone();
            if let Some(init) = owner_static_field_inits.get(&field) {
                *expr = init.clone();
            } else {
                expr.kind = ExprKind::StaticAccess {
                    class: Box::new(Expression::ident(owner_name)),
                    member: Box::new(Expression::ident(&field)),
                };
            }
        }
        ExprKind::Member { object, field, .. }
            if java_expr_dotted_name(object).as_deref() == Some(owner_name)
                && owner_static_fields.contains(field) =>
        {
            let field = field.clone();
            if let Some(init) = owner_static_field_inits.get(&field) {
                *expr = init.clone();
            } else {
                expr.kind = ExprKind::StaticAccess {
                    class: Box::new(Expression::ident(owner_name)),
                    member: Box::new(Expression::ident(&field)),
                };
            }
        }
        ExprKind::Call { callee, args, .. } => {
            rewrite_java_outer_static_refs_expr(
                callee,
                owner_name,
                owner_static_fields,
                owner_static_field_inits,
            );
            for arg in args {
                rewrite_java_outer_static_refs_expr(
                    &mut arg.value,
                    owner_name,
                    owner_static_fields,
                    owner_static_field_inits,
                );
            }
        }
        ExprKind::Member { object, .. } => {
            rewrite_java_outer_static_refs_expr(
                object,
                owner_name,
                owner_static_fields,
                owner_static_field_inits,
            );
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_java_outer_static_refs_expr(
                object,
                owner_name,
                owner_static_fields,
                owner_static_field_inits,
            );
            rewrite_java_outer_static_refs_expr(
                index,
                owner_name,
                owner_static_fields,
                owner_static_field_inits,
            );
        }
        ExprKind::Binary { left, right, .. } => {
            rewrite_java_outer_static_refs_expr(
                left,
                owner_name,
                owner_static_fields,
                owner_static_field_inits,
            );
            rewrite_java_outer_static_refs_expr(
                right,
                owner_name,
                owner_static_fields,
                owner_static_field_inits,
            );
        }
        ExprKind::Unary { expr: inner, .. } => {
            rewrite_java_outer_static_refs_expr(
                inner,
                owner_name,
                owner_static_fields,
                owner_static_field_inits,
            );
        }
        ExprKind::Assign { target, value } => {
            rewrite_java_outer_static_refs_expr(
                target,
                owner_name,
                owner_static_fields,
                owner_static_field_inits,
            );
            rewrite_java_outer_static_refs_expr(
                value,
                owner_name,
                owner_static_fields,
                owner_static_field_inits,
            );
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_java_outer_static_refs_expr(
                cond,
                owner_name,
                owner_static_fields,
                owner_static_field_inits,
            );
            rewrite_java_outer_static_refs_expr(
                then,
                owner_name,
                owner_static_fields,
                owner_static_field_inits,
            );
            rewrite_java_outer_static_refs_expr(
                else_,
                owner_name,
                owner_static_fields,
                owner_static_field_inits,
            );
        }
        ExprKind::Array(elems) => {
            for elem in elems {
                rewrite_java_outer_static_refs_expr(
                    &mut elem.value,
                    owner_name,
                    owner_static_fields,
                    owner_static_field_inits,
                );
            }
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(inner) => {
                rewrite_java_outer_static_refs_expr(
                    inner,
                    owner_name,
                    owner_static_fields,
                    owner_static_field_inits,
                );
            }
            LambdaBody::Block(stmts) => {
                rewrite_java_outer_static_refs_stmts(
                    stmts,
                    owner_name,
                    owner_static_fields,
                    owner_static_field_inits,
                );
            }
        },
        _ => {}
    }
}

fn java_prepend_outer_constructor_param(
    params: &mut Vec<Param>,
    body: &mut Vec<Statement>,
    owner_name: &str,
) {
    if !params.iter().any(|param| param.name == "__java_outer") {
        params.insert(0, java_outer_param(owner_name));
    }
    if !matches!(
        body.first().map(|stmt| &stmt.kind),
        Some(StmtKind::Assign { targets, .. })
            if matches!(
                targets.first().map(|target| &target.kind),
                Some(ExprKind::Member { field, .. }) if field == "__java_outer"
            )
    ) {
        body.insert(0, java_outer_assign_stmt());
    }
}

fn java_outer_param(owner_name: &str) -> Param {
    Param {
        name: "__java_outer".to_string(),
        type_hint: Some(owner_name.to_string()),
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    }
}

fn java_outer_assign_stmt() -> Statement {
    Statement::new(StmtKind::Assign {
        targets: vec![Expression::new(ExprKind::Member {
            object: Box::new(Expression::new(ExprKind::This)),
            field: "__java_outer".to_string(),
            null_safe: false,
        })],
        value: Expression::ident("__java_outer"),
    })
}

fn java_inner_field_init_assign_stmt(name: &str, value: Expression) -> Statement {
    Statement::new(StmtKind::Assign {
        targets: vec![Expression::new(ExprKind::Member {
            object: Box::new(Expression::new(ExprKind::This)),
            field: name.to_string(),
            null_safe: false,
        })],
        value,
    })
}

fn java_insert_inner_field_inits(body: &mut Vec<Statement>, field_inits: &[Statement]) {
    if field_inits.is_empty() {
        return;
    }
    let insert_at = if matches!(
        body.first().map(|stmt| &stmt.kind),
        Some(StmtKind::Assign { targets, .. })
            if matches!(
                targets.first().map(|target| &target.kind),
                Some(ExprKind::Member { field, .. }) if field == "__java_outer"
            )
    ) {
        1
    } else {
        0
    };
    for (offset, stmt) in field_inits.iter().cloned().enumerate() {
        body.insert(insert_at + offset, stmt);
    }
}

fn java_outer_expr() -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(Expression::new(ExprKind::This)),
        field: "__java_outer".to_string(),
        null_safe: false,
    })
}

fn java_typed_outer_expr(owner_name: &str) -> Expression {
    Expression::new(ExprKind::Cast {
        expr: Box::new(java_outer_expr()),
        type_name: owner_name.to_string(),
    })
}

fn rewrite_java_nested_type_refs_in_members(
    members: &mut [ClassMember],
    nested_types: &HashMap<String, (String, bool)>,
) {
    for member in members {
        match member {
            ClassMember::Field {
                init: Some(init), ..
            } => {
                rewrite_java_nested_type_refs_expr(init, nested_types);
            }
            ClassMember::Constructor { body, .. } => {
                rewrite_java_nested_type_refs_stmts(body, nested_types);
            }
            ClassMember::Method(stmt) => rewrite_java_nested_type_refs_stmt(stmt, nested_types),
            ClassMember::NestedType(_) => {}
            _ => {}
        }
    }
}

fn rewrite_java_nested_type_refs_stmt(
    stmt: &mut Statement,
    nested_types: &HashMap<String, (String, bool)>,
) {
    match &mut stmt.kind {
        StmtKind::FunctionDecl {
            return_type, body, ..
        } => {
            if let Some(type_hint) = return_type {
                if let Some((qualified, _)) = nested_types.get(type_hint) {
                    *type_hint = qualified.clone();
                }
            }
            rewrite_java_nested_type_refs_stmts(body, nested_types);
        }
        StmtKind::Block(body) => {
            rewrite_java_nested_type_refs_stmts(body, nested_types);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_java_nested_type_refs_expr(init, nested_types);
                }
                if let Some(type_hint) = &mut decl.type_hint {
                    if let Some((qualified, _)) = nested_types.get(type_hint) {
                        *type_hint = qualified.clone();
                    }
                }
            }
        }
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            rewrite_java_nested_type_refs_expr(expr, nested_types);
        }
        StmtKind::Assign { targets, value } => {
            rewrite_java_nested_type_refs_expr(value, nested_types);
            for target in targets {
                rewrite_java_nested_type_refs_expr(target, nested_types);
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_java_nested_type_refs_expr(cond, nested_types);
            rewrite_java_nested_type_refs_stmts(then_body, nested_types);
            for (elif_cond, elif_body) in elifs {
                rewrite_java_nested_type_refs_expr(elif_cond, nested_types);
                rewrite_java_nested_type_refs_stmts(elif_body, nested_types);
            }
            if let Some(else_body) = else_body {
                rewrite_java_nested_type_refs_stmts(else_body, nested_types);
            }
        }
        StmtKind::While { cond, body, .. } | StmtKind::DoWhile { cond, body, .. } => {
            rewrite_java_nested_type_refs_expr(cond, nested_types);
            rewrite_java_nested_type_refs_stmts(body, nested_types);
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                rewrite_java_nested_type_refs_stmt(init, nested_types);
            }
            if let Some(cond) = cond {
                rewrite_java_nested_type_refs_expr(cond, nested_types);
            }
            if let Some(update) = update {
                rewrite_java_nested_type_refs_expr(update, nested_types);
            }
            rewrite_java_nested_type_refs_stmts(body, nested_types);
        }
        StmtKind::ForIn { iter, body, .. } => {
            rewrite_java_nested_type_refs_expr(iter, nested_types);
            rewrite_java_nested_type_refs_stmts(body, nested_types);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_java_nested_type_refs_expr(target, nested_types);
            rewrite_java_nested_type_refs_expr(value, nested_types);
        }
        StmtKind::Throw {
            expr: Some(expr), ..
        } => {
            rewrite_java_nested_type_refs_expr(expr, nested_types);
        }
        _ => {}
    }
}

fn rewrite_java_nested_type_refs_stmts(
    stmts: &mut [Statement],
    nested_types: &HashMap<String, (String, bool)>,
) {
    for stmt in stmts {
        rewrite_java_nested_type_refs_stmt(stmt, nested_types);
    }
}

fn rewrite_java_nested_type_refs_expr(
    expr: &mut Expression,
    nested_types: &HashMap<String, (String, bool)>,
) {
    match &mut expr.kind {
        ExprKind::New { class, args } => {
            if let ExprKind::Ident(name) = &class.kind {
                if let Some((qualified, is_static)) = nested_types.get(name) {
                    class.kind = ExprKind::Ident(qualified.clone());
                    if !*is_static {
                        args.insert(0, Argument::positional(Expression::new(ExprKind::This)));
                    }
                }
            }
            for arg in args {
                rewrite_java_nested_type_refs_expr(&mut arg.value, nested_types);
            }
        }
        ExprKind::Call { callee, args, .. } => {
            if let ExprKind::Member { object, field, .. } = &mut callee.kind {
                if let Some(type_name) = java_registered_nested_type_name(object) {
                    callee.kind = ExprKind::StaticAccess {
                        class: Box::new(Expression::ident(&type_name)),
                        member: Box::new(Expression::ident(field)),
                    };
                }
            }
            rewrite_java_nested_type_refs_expr(callee, nested_types);
            for arg in args {
                rewrite_java_nested_type_refs_expr(&mut arg.value, nested_types);
            }
        }
        ExprKind::Member { object, .. } => {
            if let Some(type_name) = java_registered_nested_type_name(object) {
                let ExprKind::Member { field, .. } = &expr.kind else {
                    return;
                };
                expr.kind = ExprKind::StaticAccess {
                    class: Box::new(Expression::ident(&type_name)),
                    member: Box::new(Expression::ident(field)),
                };
                return;
            }
            rewrite_java_nested_type_refs_expr(object, nested_types);
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_java_nested_type_refs_expr(object, nested_types);
            rewrite_java_nested_type_refs_expr(index, nested_types);
        }
        ExprKind::Binary { left, right, .. } => {
            rewrite_java_nested_type_refs_expr(left, nested_types);
            rewrite_java_nested_type_refs_expr(right, nested_types);
        }
        ExprKind::Unary { expr: inner, .. } => {
            rewrite_java_nested_type_refs_expr(inner, nested_types)
        }
        ExprKind::Assign { target, value } => {
            rewrite_java_nested_type_refs_expr(target, nested_types);
            rewrite_java_nested_type_refs_expr(value, nested_types);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_java_nested_type_refs_expr(cond, nested_types);
            rewrite_java_nested_type_refs_expr(then, nested_types);
            rewrite_java_nested_type_refs_expr(else_, nested_types);
        }
        ExprKind::Array(elems) => {
            for elem in elems {
                rewrite_java_nested_type_refs_expr(&mut elem.value, nested_types);
            }
        }
        _ => {}
    }
}

fn java_registered_nested_type_name(expr: &Expression) -> Option<String> {
    let mut parts = Vec::new();
    collect_member_chain(expr, &mut parts)?;
    if parts.len() < 2 {
        return None;
    }
    let dotted = parts.join(".");
    JAVA_NESTED_TYPE_NAMES.with(|names| {
        if names.borrow().contains(&dotted) {
            Some(dotted)
        } else {
            None
        }
    })
}

fn rewrite_java_inner_outer_refs_stmt(
    stmt: &mut Statement,
    owner_name: &str,
    owner_fields: &HashSet<String>,
    owner_methods: &HashSet<String>,
) {
    match &mut stmt.kind {
        StmtKind::FunctionDecl { body, .. } | StmtKind::Block(body) => {
            rewrite_java_inner_outer_refs_stmts(body, owner_name, owner_fields, owner_methods);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_java_inner_outer_refs_expr(
                        init,
                        owner_name,
                        owner_fields,
                        owner_methods,
                    );
                }
            }
        }
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            rewrite_java_inner_outer_refs_expr(expr, owner_name, owner_fields, owner_methods);
        }
        StmtKind::Assign { targets, value } => {
            rewrite_java_inner_outer_refs_expr(value, owner_name, owner_fields, owner_methods);
            for target in targets {
                rewrite_java_inner_outer_refs_expr(target, owner_name, owner_fields, owner_methods);
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_java_inner_outer_refs_expr(cond, owner_name, owner_fields, owner_methods);
            rewrite_java_inner_outer_refs_stmts(then_body, owner_name, owner_fields, owner_methods);
            for (elif_cond, elif_body) in elifs {
                rewrite_java_inner_outer_refs_expr(
                    elif_cond,
                    owner_name,
                    owner_fields,
                    owner_methods,
                );
                rewrite_java_inner_outer_refs_stmts(
                    elif_body,
                    owner_name,
                    owner_fields,
                    owner_methods,
                );
            }
            if let Some(else_body) = else_body {
                rewrite_java_inner_outer_refs_stmts(
                    else_body,
                    owner_name,
                    owner_fields,
                    owner_methods,
                );
            }
        }
        StmtKind::While { cond, body, .. } => {
            rewrite_java_inner_outer_refs_expr(cond, owner_name, owner_fields, owner_methods);
            rewrite_java_inner_outer_refs_stmts(body, owner_name, owner_fields, owner_methods);
        }
        _ => {}
    }
}

fn rewrite_java_inner_outer_refs_stmts(
    stmts: &mut [Statement],
    owner_name: &str,
    owner_fields: &HashSet<String>,
    owner_methods: &HashSet<String>,
) {
    for stmt in stmts {
        rewrite_java_inner_outer_refs_stmt(stmt, owner_name, owner_fields, owner_methods);
    }
}

fn rewrite_java_inner_outer_refs_expr(
    expr: &mut Expression,
    owner_name: &str,
    owner_fields: &HashSet<String>,
    owner_methods: &HashSet<String>,
) {
    match &mut expr.kind {
        ExprKind::Ident(name) if owner_fields.contains(name) => {
            let field = name.clone();
            expr.kind = ExprKind::Member {
                object: Box::new(java_typed_outer_expr(owner_name)),
                field,
                null_safe: false,
            };
        }
        ExprKind::Call { callee, args, .. } => {
            for arg in &mut *args {
                rewrite_java_inner_outer_refs_expr(
                    &mut arg.value,
                    owner_name,
                    owner_fields,
                    owner_methods,
                );
            }
            if let ExprKind::Ident(name) = &callee.kind {
                if owner_methods.contains(name) {
                    let method = name.clone();
                    callee.kind = ExprKind::Member {
                        object: Box::new(java_typed_outer_expr(owner_name)),
                        field: method,
                        null_safe: false,
                    };
                    return;
                }
            }
            rewrite_java_inner_outer_refs_expr(callee, owner_name, owner_fields, owner_methods);
        }
        ExprKind::Member { object, field, .. } => {
            if field == "this" && java_expr_dotted_name(object).as_deref() == Some(owner_name) {
                expr.kind = java_typed_outer_expr(owner_name).kind;
                return;
            }
            rewrite_java_inner_outer_refs_expr(object, owner_name, owner_fields, owner_methods);
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_java_inner_outer_refs_expr(object, owner_name, owner_fields, owner_methods);
            rewrite_java_inner_outer_refs_expr(index, owner_name, owner_fields, owner_methods);
        }
        ExprKind::Binary { left, right, .. } => {
            rewrite_java_inner_outer_refs_expr(left, owner_name, owner_fields, owner_methods);
            rewrite_java_inner_outer_refs_expr(right, owner_name, owner_fields, owner_methods);
        }
        ExprKind::Unary { expr: inner, .. } => {
            rewrite_java_inner_outer_refs_expr(inner, owner_name, owner_fields, owner_methods);
        }
        ExprKind::Assign { target, value } => {
            rewrite_java_inner_outer_refs_expr(target, owner_name, owner_fields, owner_methods);
            rewrite_java_inner_outer_refs_expr(value, owner_name, owner_fields, owner_methods);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_java_inner_outer_refs_expr(cond, owner_name, owner_fields, owner_methods);
            rewrite_java_inner_outer_refs_expr(then, owner_name, owner_fields, owner_methods);
            rewrite_java_inner_outer_refs_expr(else_, owner_name, owner_fields, owner_methods);
        }
        ExprKind::Array(elems) => {
            for elem in elems {
                rewrite_java_inner_outer_refs_expr(
                    &mut elem.value,
                    owner_name,
                    owner_fields,
                    owner_methods,
                );
            }
        }
        _ => {}
    }
}

fn java_expr_reads_java_outer(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => name == "__java_outer",
        ExprKind::Member { object, field, .. } => {
            field == "__java_outer" || java_expr_reads_java_outer(object)
        }
        ExprKind::Call { callee, args, .. } => {
            java_expr_reads_java_outer(callee)
                || args
                    .iter()
                    .any(|arg| java_expr_reads_java_outer(&arg.value))
        }
        ExprKind::Index { object, index, .. } => {
            java_expr_reads_java_outer(object) || java_expr_reads_java_outer(index)
        }
        ExprKind::Binary { left, right, .. } => {
            java_expr_reads_java_outer(left) || java_expr_reads_java_outer(right)
        }
        ExprKind::Unary { expr, .. } | ExprKind::Cast { expr, .. } => {
            java_expr_reads_java_outer(expr)
        }
        ExprKind::Assign { target, value } => {
            java_expr_reads_java_outer(target) || java_expr_reads_java_outer(value)
        }
        ExprKind::Ternary { cond, then, else_ } => {
            java_expr_reads_java_outer(cond)
                || java_expr_reads_java_outer(then)
                || java_expr_reads_java_outer(else_)
        }
        ExprKind::Array(elems) => elems
            .iter()
            .any(|elem| java_expr_reads_java_outer(&elem.value)),
        ExprKind::New { class, args } => {
            java_expr_reads_java_outer(class)
                || args
                    .iter()
                    .any(|arg| java_expr_reads_java_outer(&arg.value))
        }
        ExprKind::StaticAccess { class, member } => {
            java_expr_reads_java_outer(class) || java_expr_reads_java_outer(member)
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(expr) => java_expr_reads_java_outer(expr),
            LambdaBody::Block(stmts) => stmts.iter().any(java_stmt_reads_java_outer),
        },
        _ => false,
    }
}

fn java_stmt_reads_java_outer(stmt: &Statement) -> bool {
    match &stmt.kind {
        StmtKind::FunctionDecl { body, .. } | StmtKind::Block(body) => {
            body.iter().any(java_stmt_reads_java_outer)
        }
        StmtKind::VarDecl { declarations, .. } => declarations
            .iter()
            .any(|decl| decl.init.as_ref().is_some_and(java_expr_reads_java_outer)),
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => java_expr_reads_java_outer(expr),
        StmtKind::Assign { targets, value } => {
            java_expr_reads_java_outer(value) || targets.iter().any(java_expr_reads_java_outer)
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            java_expr_reads_java_outer(target) || java_expr_reads_java_outer(value)
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            java_expr_reads_java_outer(cond)
                || then_body.iter().any(java_stmt_reads_java_outer)
                || elifs.iter().any(|(cond, body)| {
                    java_expr_reads_java_outer(cond) || body.iter().any(java_stmt_reads_java_outer)
                })
                || else_body
                    .as_ref()
                    .is_some_and(|body| body.iter().any(java_stmt_reads_java_outer))
        }
        StmtKind::While { cond, body, .. } | StmtKind::DoWhile { cond, body, .. } => {
            java_expr_reads_java_outer(cond) || body.iter().any(java_stmt_reads_java_outer)
        }
        _ => false,
    }
}

fn rewrite_java_user_tostring_calls(body: &mut [Statement]) {
    let mut tostring_classes = HashSet::new();
    collect_java_tostring_classes(body, &mut tostring_classes);
    let mut enum_values = HashMap::new();
    collect_java_enum_values(body, &mut enum_values);
    // Enums now walk to ClassDecl — the authoritative member map is the
    // walker's own registry.
    JAVA_ENUM_VALUES.with(|values| {
        for (enum_name, members) in values.borrow().iter() {
            enum_values
                .entry(enum_name.clone())
                .or_insert_with(|| members.clone());
        }
    });
    let mut double_fields = HashSet::new();
    collect_java_double_fields(body, &mut double_fields);
    let mut double_methods = HashSet::new();
    collect_java_double_methods(body, &mut double_methods);
    rewrite_java_double_field_print_tree(body, &double_fields);
    rewrite_java_double_method_print_tree(body, &double_methods);
    rewrite_java_tostring_stmts(
        body,
        &tostring_classes,
        &enum_values,
        None,
        &mut HashMap::new(),
    );
}

fn collect_java_enum_values(stmts: &[Statement], out: &mut HashMap<String, Vec<String>>) {
    for stmt in stmts {
        match &stmt.kind {
            StmtKind::EnumDecl { name, members, .. } => {
                out.insert(
                    name.clone(),
                    members.iter().map(|member| member.name.clone()).collect(),
                );
            }
            StmtKind::ClassDecl { members, .. } => {
                for member in members {
                    if let ClassMember::NestedType(nested) = member {
                        collect_java_enum_values(std::slice::from_ref(nested), out);
                    }
                }
            }
            StmtKind::Block(stmts) => collect_java_enum_values(stmts, out),
            _ => {}
        }
    }
}

fn collect_java_double_fields(stmts: &[Statement], out: &mut HashSet<String>) {
    for stmt in stmts {
        if let StmtKind::ClassDecl { members, .. } = &stmt.kind {
            for member in members {
                if let ClassMember::Field {
                    name,
                    type_hint: Some(type_hint),
                    ..
                } = member
                {
                    if matches!(type_hint.as_str(), "double" | "Double") {
                        out.insert(name.clone());
                    }
                }
            }
        }
    }
}

fn collect_java_double_methods(stmts: &[Statement], out: &mut HashSet<String>) {
    for stmt in stmts {
        match &stmt.kind {
            StmtKind::ClassDecl { members, .. } => {
                for member in members {
                    match member {
                        ClassMember::Method(method) => {
                            if let StmtKind::FunctionDecl {
                                name,
                                return_type: Some(return_type),
                                ..
                            } = &method.kind
                            {
                                if matches!(return_type.as_str(), "double" | "Double") {
                                    out.insert(name.clone());
                                }
                            }
                        }
                        ClassMember::NestedType(nested) => {
                            collect_java_double_methods(std::slice::from_ref(nested), out);
                        }
                        _ => {}
                    }
                }
            }
            StmtKind::Block(stmts) => collect_java_double_methods(stmts, out),
            _ => {}
        }
    }
}

fn rewrite_java_double_field_print_tree(stmts: &mut [Statement], double_fields: &HashSet<String>) {
    for stmt in stmts {
        rewrite_java_double_field_prints(std::slice::from_mut(stmt), double_fields);
        if let StmtKind::ClassDecl { members, .. } = &mut stmt.kind {
            for member in members {
                match member {
                    ClassMember::Constructor { body, .. } => {
                        rewrite_java_double_field_print_tree(body, double_fields);
                    }
                    ClassMember::Method(method) => {
                        if let StmtKind::FunctionDecl { body, .. } = &mut method.kind {
                            rewrite_java_double_field_print_tree(body, double_fields);
                        }
                    }
                    _ => {}
                }
            }
        }
    }
}

fn rewrite_java_double_method_print_tree(
    stmts: &mut [Statement],
    double_methods: &HashSet<String>,
) {
    for stmt in stmts {
        rewrite_java_double_method_prints(std::slice::from_mut(stmt), double_methods);
        if let StmtKind::ClassDecl { members, .. } = &mut stmt.kind {
            for member in members {
                match member {
                    ClassMember::Constructor { body, .. } => {
                        rewrite_java_double_method_print_tree(body, double_methods);
                    }
                    ClassMember::Method(method) => {
                        if let StmtKind::FunctionDecl { body, .. } = &mut method.kind {
                            rewrite_java_double_method_print_tree(body, double_methods);
                        }
                    }
                    _ => {}
                }
            }
        }
    }
}

fn collect_java_tostring_classes(body: &[Statement], out: &mut HashSet<String>) {
    for stmt in body {
        match &stmt.kind {
            StmtKind::ClassDecl { name, members, .. } => {
                if members.iter().any(|member| {
                    matches!(
                        member,
                        ClassMember::Method(method)
                            if matches!(&method.kind, StmtKind::FunctionDecl { name, .. } if name == "toString")
                    )
                }) {
                    out.insert(name.clone());
                }
                for member in members {
                    if let ClassMember::NestedType(nested) = member {
                        collect_java_tostring_classes(std::slice::from_ref(nested), out);
                    }
                }
            }
            // Every enum has toString (JLS §8.9.3): user-declared or the
            // walker-synthesized `return this.__name`.
            StmtKind::EnumDecl { name, .. } => {
                out.insert(name.clone());
            }
            _ => {}
        }
    }
}

fn rewrite_java_tostring_stmts(
    stmts: &mut [Statement],
    tostring_classes: &HashSet<String>,
    enum_values: &HashMap<String, Vec<String>>,
    current_class: Option<&str>,
    locals: &mut HashMap<String, String>,
) {
    for stmt in stmts {
        match &mut stmt.kind {
            StmtKind::ClassDecl { name, members, .. } => {
                let double_fields: std::collections::HashSet<String> = members
                    .iter()
                    .filter_map(|member| match member {
                        ClassMember::Field {
                            name,
                            type_hint: Some(type_hint),
                            ..
                        } if matches!(type_hint.as_str(), "double" | "Double") => {
                            Some(name.clone())
                        }
                        _ => None,
                    })
                    .collect();
                let field_types: HashMap<String, String> = members
                    .iter()
                    .filter_map(|member| match member {
                        ClassMember::Field {
                            name, type_hint, ..
                        } => type_hint.as_ref().map(|t| (name.clone(), t.clone())),
                        _ => None,
                    })
                    .collect();
                let method_return_types: HashMap<String, String> = members
                    .iter()
                    .filter_map(|member| match member {
                        ClassMember::Method(method) => match &method.kind {
                            StmtKind::FunctionDecl {
                                name,
                                params,
                                return_type: Some(return_type),
                                ..
                            } if params.is_empty() => {
                                Some((format!("{name}()"), return_type.clone()))
                            }
                            _ => None,
                        },
                        _ => None,
                    })
                    .collect();
                JAVA_STATIC_FIELD_VARS.with(|vars| {
                    let mut vars = vars.borrow_mut();
                    for member in members.iter() {
                        if let ClassMember::Field {
                            name: field_name,
                            modifiers,
                            type_hint,
                            ..
                        } = member
                        {
                            if modifiers.is_static {
                                vars.insert(format!("{name}.{field_name}"));
                                if let Some(type_hint) = type_hint {
                                    JAVA_STATIC_FIELD_TYPES.with(|types| {
                                        types.borrow_mut().insert(
                                            format!("{name}.{field_name}"),
                                            type_hint.clone(),
                                        );
                                    });
                                }
                            }
                        }
                    }
                });
                for member in members {
                    match member {
                        ClassMember::Constructor { params, body, .. } => {
                            let mut local_types: HashMap<String, String> = params
                                .iter()
                                .filter_map(|p| {
                                    p.type_hint.as_ref().map(|t| (p.name.clone(), t.clone()))
                                })
                                .collect();
                            local_types.extend(
                                field_types
                                    .iter()
                                    .map(|(name, ty)| (name.clone(), ty.clone())),
                            );
                            local_types.extend(
                                method_return_types
                                    .iter()
                                    .map(|(name, ty)| (name.clone(), ty.clone())),
                            );
                            rewrite_java_tostring_stmts(
                                body,
                                tostring_classes,
                                enum_values,
                                Some(name),
                                &mut local_types,
                            );
                        }
                        ClassMember::Method(method) => {
                            if let StmtKind::FunctionDecl { params, body, .. } = &mut method.kind {
                                rewrite_java_double_field_prints(body, &double_fields);
                                let mut local_types: HashMap<String, String> = params
                                    .iter()
                                    .filter_map(|p| {
                                        p.type_hint.as_ref().map(|t| (p.name.clone(), t.clone()))
                                    })
                                    .collect();
                                local_types.extend(
                                    field_types
                                        .iter()
                                        .map(|(name, ty)| (name.clone(), ty.clone())),
                                );
                                local_types.extend(
                                    method_return_types
                                        .iter()
                                        .map(|(name, ty)| (name.clone(), ty.clone())),
                                );
                                rewrite_java_tostring_stmts(
                                    body,
                                    tostring_classes,
                                    enum_values,
                                    Some(name),
                                    &mut local_types,
                                );
                            }
                        }
                        ClassMember::NestedType(nested) => {
                            rewrite_java_tostring_stmts(
                                std::slice::from_mut(nested),
                                tostring_classes,
                                enum_values,
                                current_class,
                                &mut locals.clone(),
                            );
                        }
                        _ => {}
                    }
                }
            }
            StmtKind::VarDecl { declarations, .. } => {
                for decl in declarations {
                    if let Some(init) = &mut decl.init {
                        rewrite_java_tostring_expr(
                            init,
                            tostring_classes,
                            enum_values,
                            current_class,
                            locals,
                        );
                    }
                    if let (BindingPattern::Ident(name), Some(init)) = (&decl.pattern, &decl.init) {
                        if name.starts_with("__java_switch_value_") {
                            if let ExprKind::Ident(source_name) = &init.kind {
                                if let Some(type_hint) = locals.get(source_name).cloned() {
                                    locals.insert(name.clone(), type_hint);
                                }
                            }
                        }
                    }
                    if let (BindingPattern::Ident(name), Some(type_hint)) =
                        (&decl.pattern, &decl.type_hint)
                    {
                        locals.insert(name.clone(), type_hint.clone());
                    }
                }
            }
            StmtKind::Assign { targets, value } => {
                rewrite_java_tostring_expr(
                    value,
                    tostring_classes,
                    enum_values,
                    current_class,
                    locals,
                );
                for target in &mut *targets {
                    rewrite_java_tostring_expr(
                        target,
                        tostring_classes,
                        enum_values,
                        current_class,
                        locals,
                    );
                }
            }
            StmtKind::CompoundAssign { target, value, .. } => {
                rewrite_java_tostring_expr(
                    value,
                    tostring_classes,
                    enum_values,
                    current_class,
                    locals,
                );
                rewrite_java_tostring_expr(
                    target,
                    tostring_classes,
                    enum_values,
                    current_class,
                    locals,
                );
            }
            StmtKind::Expr(expr)
            | StmtKind::Return(Some(expr))
            | StmtKind::Throw {
                expr: Some(expr), ..
            } => {
                rewrite_java_tostring_expr(
                    expr,
                    tostring_classes,
                    enum_values,
                    current_class,
                    locals,
                );
                if matches!(stmt.kind, StmtKind::Expr(_)) {
                    rewrite_java_map_for_each_stmt(stmt);
                }
            }
            StmtKind::If {
                cond,
                then_body,
                elifs,
                else_body,
            } => {
                rewrite_java_tostring_expr(
                    cond,
                    tostring_classes,
                    enum_values,
                    current_class,
                    locals,
                );
                rewrite_java_tostring_stmts(
                    then_body,
                    tostring_classes,
                    enum_values,
                    current_class,
                    &mut locals.clone(),
                );
                for (elif_cond, elif_body) in elifs {
                    rewrite_java_tostring_expr(
                        elif_cond,
                        tostring_classes,
                        enum_values,
                        current_class,
                        locals,
                    );
                    rewrite_java_tostring_stmts(
                        elif_body,
                        tostring_classes,
                        enum_values,
                        current_class,
                        &mut locals.clone(),
                    );
                }
                if let Some(else_body) = else_body {
                    rewrite_java_tostring_stmts(
                        else_body,
                        tostring_classes,
                        enum_values,
                        current_class,
                        &mut locals.clone(),
                    );
                }
            }
            StmtKind::While { cond, body, .. } => {
                rewrite_java_tostring_expr(
                    cond,
                    tostring_classes,
                    enum_values,
                    current_class,
                    locals,
                );
                rewrite_java_tostring_stmts(
                    body,
                    tostring_classes,
                    enum_values,
                    current_class,
                    &mut locals.clone(),
                );
            }
            StmtKind::ForIn {
                var,
                key,
                iter,
                body,
                else_body,
                ..
            } => {
                rewrite_java_tostring_expr(
                    iter,
                    tostring_classes,
                    enum_values,
                    current_class,
                    locals,
                );
                let mut loop_locals = locals.clone();
                loop_locals.insert(var.clone(), "Object".to_string());
                if let Some(key) = key {
                    loop_locals.insert(key.clone(), "Object".to_string());
                }
                rewrite_java_tostring_stmts(
                    body,
                    tostring_classes,
                    enum_values,
                    current_class,
                    &mut loop_locals,
                );
                if let Some(else_body) = else_body {
                    rewrite_java_tostring_stmts(
                        else_body,
                        tostring_classes,
                        enum_values,
                        current_class,
                        &mut locals.clone(),
                    );
                }
            }
            StmtKind::Block(body) => {
                rewrite_java_tostring_stmts(
                    body,
                    tostring_classes,
                    enum_values,
                    current_class,
                    &mut locals.clone(),
                );
            }
            StmtKind::Try {
                body,
                catches,
                else_body,
                finally,
            } => {
                rewrite_java_tostring_stmts(
                    body,
                    tostring_classes,
                    enum_values,
                    current_class,
                    locals,
                );
                for c in catches {
                    let mut catch_locals = locals.clone();
                    if let (Some(var_name), Some(ty)) = (&c.var_name, c.types.first()) {
                        catch_locals.insert(var_name.clone(), ty.clone());
                    }
                    rewrite_java_tostring_stmts(
                        &mut c.body,
                        tostring_classes,
                        enum_values,
                        current_class,
                        &mut catch_locals,
                    );
                }
                if let Some(else_body) = else_body {
                    rewrite_java_tostring_stmts(
                        else_body,
                        tostring_classes,
                        enum_values,
                        current_class,
                        locals,
                    );
                }
                if let Some(finally) = finally {
                    rewrite_java_tostring_stmts(
                        finally,
                        tostring_classes,
                        enum_values,
                        current_class,
                        locals,
                    );
                }
            }
            StmtKind::For {
                init,
                cond,
                update,
                body,
            } => {
                if let Some(init) = init {
                    rewrite_java_tostring_stmts(
                        std::slice::from_mut(init),
                        tostring_classes,
                        enum_values,
                        current_class,
                        locals,
                    );
                }
                if let Some(cond) = cond {
                    rewrite_java_tostring_expr(
                        cond,
                        tostring_classes,
                        enum_values,
                        current_class,
                        locals,
                    );
                }
                if let Some(update) = update {
                    rewrite_java_tostring_expr(
                        update,
                        tostring_classes,
                        enum_values,
                        current_class,
                        locals,
                    );
                }
                rewrite_java_tostring_stmts(
                    body,
                    tostring_classes,
                    enum_values,
                    current_class,
                    locals,
                );
            }
            _ => {}
        }
    }
}

fn rewrite_java_map_for_each_stmt(stmt: &mut Statement) {
    let span = stmt.span;
    let StmtKind::Expr(expr) = &stmt.kind else {
        return;
    };
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return;
    };
    if !matches!(callee.kind, ExprKind::Ident(ref name) if name == "__java_map_for_each")
        || args.len() != 2
    {
        return;
    }
    let ExprKind::Lambda { params, body, .. } = &args[1].value.kind else {
        return;
    };
    if params.len() != 2 {
        return;
    }

    let key_name = params[0].name.clone();
    let value_name = params[1].name.clone();
    let entry_name = format!(
        "__java_map_for_each_entry_{}_{}",
        span.start_line, span.start_col
    );
    let mut body_stmts = match body {
        LambdaBody::Expr(inner) => vec![Statement::with_span(
            StmtKind::Expr((**inner).clone()),
            span,
        )],
        LambdaBody::Block(stmts) => stmts.clone(),
    };
    for body_stmt in &mut body_stmts {
        substitute_java_map_for_each_params_stmt(body_stmt, &key_name, &value_name, &entry_name);
    }

    let iter = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("__java_map_entry_set")),
        args: vec![Argument::positional(args[0].value.clone())],
        optional: false,
    });
    stmt.kind = StmtKind::ForIn {
        var: entry_name,
        key: None,
        iter,
        body: body_stmts,
        of: true,
        else_body: None,
        is_async: false,
    };
}

fn java_map_for_each_entry_expr(entry_name: &str, index: i64) -> Expression {
    Expression::new(ExprKind::Index {
        object: Box::new(Expression::ident(entry_name)),
        index: Box::new(Expression::new(ExprKind::Lit(Literal::Int(index)))),
        null_safe: false,
    })
}

fn substitute_java_map_for_each_params_stmt(
    stmt: &mut Statement,
    key_name: &str,
    value_name: &str,
    entry_name: &str,
) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            substitute_java_map_for_each_params_expr(expr, key_name, value_name, entry_name);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    substitute_java_map_for_each_params_expr(
                        init, key_name, value_name, entry_name,
                    );
                }
            }
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                substitute_java_map_for_each_params_expr(target, key_name, value_name, entry_name);
            }
            substitute_java_map_for_each_params_expr(value, key_name, value_name, entry_name);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            substitute_java_map_for_each_params_expr(target, key_name, value_name, entry_name);
            substitute_java_map_for_each_params_expr(value, key_name, value_name, entry_name);
        }
        StmtKind::Block(stmts) => {
            for nested in stmts {
                substitute_java_map_for_each_params_stmt(nested, key_name, value_name, entry_name);
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            substitute_java_map_for_each_params_expr(cond, key_name, value_name, entry_name);
            for nested in then_body {
                substitute_java_map_for_each_params_stmt(nested, key_name, value_name, entry_name);
            }
            for (elif_cond, elif_body) in elifs {
                substitute_java_map_for_each_params_expr(
                    elif_cond, key_name, value_name, entry_name,
                );
                for nested in elif_body {
                    substitute_java_map_for_each_params_stmt(
                        nested, key_name, value_name, entry_name,
                    );
                }
            }
            if let Some(else_body) = else_body {
                for nested in else_body {
                    substitute_java_map_for_each_params_stmt(
                        nested, key_name, value_name, entry_name,
                    );
                }
            }
        }
        StmtKind::While { cond, body, .. } | StmtKind::DoWhile { cond, body, .. } => {
            substitute_java_map_for_each_params_expr(cond, key_name, value_name, entry_name);
            for nested in body {
                substitute_java_map_for_each_params_stmt(nested, key_name, value_name, entry_name);
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                substitute_java_map_for_each_params_stmt(init, key_name, value_name, entry_name);
            }
            if let Some(cond) = cond {
                substitute_java_map_for_each_params_expr(cond, key_name, value_name, entry_name);
            }
            if let Some(update) = update {
                substitute_java_map_for_each_params_expr(update, key_name, value_name, entry_name);
            }
            for nested in body {
                substitute_java_map_for_each_params_stmt(nested, key_name, value_name, entry_name);
            }
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            substitute_java_map_for_each_params_expr(iter, key_name, value_name, entry_name);
            for nested in body {
                substitute_java_map_for_each_params_stmt(nested, key_name, value_name, entry_name);
            }
            if let Some(else_body) = else_body {
                for nested in else_body {
                    substitute_java_map_for_each_params_stmt(
                        nested, key_name, value_name, entry_name,
                    );
                }
            }
        }
        StmtKind::Throw {
            expr: Some(expr), ..
        } => substitute_java_map_for_each_params_expr(expr, key_name, value_name, entry_name),
        _ => {}
    }
}

fn substitute_java_map_for_each_params_expr(
    expr: &mut Expression,
    key_name: &str,
    value_name: &str,
    entry_name: &str,
) {
    match &mut expr.kind {
        ExprKind::Ident(name) if name == key_name => {
            *expr = java_map_for_each_entry_expr(entry_name, 0);
        }
        ExprKind::Ident(name) if name == value_name => {
            *expr = java_map_for_each_entry_expr(entry_name, 1);
        }
        ExprKind::Binary { left, right, .. } => {
            substitute_java_map_for_each_params_expr(left, key_name, value_name, entry_name);
            substitute_java_map_for_each_params_expr(right, key_name, value_name, entry_name);
        }
        ExprKind::Unary { expr: inner, .. }
        | ExprKind::Spread(inner)
        | ExprKind::Await(inner)
        | ExprKind::YieldFrom(inner)
        | ExprKind::Void(inner)
        | ExprKind::Delete(inner)
        | ExprKind::TypeOf(inner)
        | ExprKind::RefLoad(inner) => {
            substitute_java_map_for_each_params_expr(inner, key_name, value_name, entry_name);
        }
        ExprKind::Yield(Some(inner)) => {
            substitute_java_map_for_each_params_expr(inner, key_name, value_name, entry_name);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            substitute_java_map_for_each_params_expr(cond, key_name, value_name, entry_name);
            substitute_java_map_for_each_params_expr(then, key_name, value_name, entry_name);
            substitute_java_map_for_each_params_expr(else_, key_name, value_name, entry_name);
        }
        ExprKind::Member { object, .. } => {
            substitute_java_map_for_each_params_expr(object, key_name, value_name, entry_name);
        }
        ExprKind::Index { object, index, .. } => {
            substitute_java_map_for_each_params_expr(object, key_name, value_name, entry_name);
            substitute_java_map_for_each_params_expr(index, key_name, value_name, entry_name);
        }
        ExprKind::Call { callee, args, .. } => {
            substitute_java_map_for_each_params_expr(callee, key_name, value_name, entry_name);
            for arg in args {
                substitute_java_map_for_each_params_expr(
                    &mut arg.value,
                    key_name,
                    value_name,
                    entry_name,
                );
            }
        }
        ExprKind::Assign { target, value } => {
            substitute_java_map_for_each_params_expr(target, key_name, value_name, entry_name);
            substitute_java_map_for_each_params_expr(value, key_name, value_name, entry_name);
        }
        ExprKind::Array(elems) => {
            for elem in elems {
                substitute_java_map_for_each_params_expr(
                    &mut elem.value,
                    key_name,
                    value_name,
                    entry_name,
                );
            }
        }
        ExprKind::Tuple(elems) | ExprKind::Set(elems) | ExprKind::Sequence(elems) => {
            for elem in elems {
                substitute_java_map_for_each_params_expr(elem, key_name, value_name, entry_name);
            }
        }
        ExprKind::New { class, args } => {
            substitute_java_map_for_each_params_expr(class, key_name, value_name, entry_name);
            for arg in args {
                substitute_java_map_for_each_params_expr(
                    &mut arg.value,
                    key_name,
                    value_name,
                    entry_name,
                );
            }
        }
        ExprKind::Lambda { .. } => {}
        _ => {}
    }
}

fn rewrite_java_double_field_prints(stmts: &mut [Statement], double_fields: &HashSet<String>) {
    for stmt in stmts {
        let StmtKind::Expr(expr) = &mut stmt.kind else {
            continue;
        };
        let ExprKind::Call { callee, args, .. } = &mut expr.kind else {
            continue;
        };
        if !matches!(callee.kind, ExprKind::Ident(ref name) if name == "println" || name == "print")
            || args.len() != 1
        {
            continue;
        }
        let ExprKind::Member { field, .. } = &args[0].value.kind else {
            continue;
        };
        if !double_fields.contains(field) {
            continue;
        }
        let value = args[0].value.clone();
        args[0].value = Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__java_double_to_string")),
            args: vec![Argument::positional(value)],
            optional: false,
        });
    }
}

fn rewrite_java_double_method_prints(stmts: &mut [Statement], double_methods: &HashSet<String>) {
    for stmt in stmts {
        let StmtKind::Expr(expr) = &mut stmt.kind else {
            continue;
        };
        let ExprKind::Call { callee, args, .. } = &mut expr.kind else {
            continue;
        };
        if !matches!(callee.kind, ExprKind::Ident(ref name) if name == "println" || name == "print")
            || args.len() != 1
        {
            continue;
        }
        if java_is_double_print_call(&args[0].value) {
            let value = args[0].value.clone();
            args[0].value = Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__java_double_to_string")),
                args: vec![Argument::positional(value)],
                optional: false,
            });
            continue;
        }
        let ExprKind::Call { callee: inner, .. } = &args[0].value.kind else {
            continue;
        };
        let ExprKind::Member { field, .. } = &inner.kind else {
            continue;
        };
        if !double_methods.contains(field) {
            continue;
        }
        let value = args[0].value.clone();
        args[0].value = Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__java_double_to_string")),
            args: vec![Argument::positional(value)],
            optional: false,
        });
    }
}

fn java_is_double_print_call(expr: &Expression) -> bool {
    let ExprKind::Call { callee, .. } = &expr.kind else {
        return false;
    };
    matches!(
        &callee.kind,
        ExprKind::Ident(name) if matches!(name.as_str(), "Math.copySign" | "StrictMath.copySign")
    )
}

fn rewrite_java_tostring_expr(
    expr: &mut Expression,
    tostring_classes: &HashSet<String>,
    enum_values: &HashMap<String, Vec<String>>,
    current_class: Option<&str>,
    locals: &HashMap<String, String>,
) {
    if let Some(replacement) = java_bigint_constant_replacement(expr) {
        *expr = replacement;
        return;
    }
    if let Some(replacement) = java_bigdecimal_constant_replacement(expr) {
        *expr = replacement;
        return;
    }

    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            for arg in &mut *args {
                rewrite_java_tostring_expr(
                    &mut arg.value,
                    tostring_classes,
                    enum_values,
                    current_class,
                    locals,
                );
            }
            if let Some(rewritten) = rewrite_java_enum_set_static_call(callee, args, enum_values) {
                *expr = rewritten;
                return;
            }
            // `EnumType.valueOf(x)` → the walker-synthesized static (the
            // `valueOf` name is intercepted before user-class static dispatch).
            if let ExprKind::Member { object, field, .. } = &mut callee.kind {
                if field == "valueOf" {
                    if let ExprKind::Ident(type_name) = &object.kind {
                        if enum_values.contains_key(type_name) {
                            *field = "__j_enum_value_of".to_string();
                        }
                    }
                }
            }
            // println(Object) prints String.valueOf(x) → x.toString()
            // (java.io.PrintStream). Wrap args whose static type is a user
            // class with toString or an enum constant (JLS §8.9.3 toString).
            if let ExprKind::Ident(fn_name) = &callee.kind {
                if matches!(
                    fn_name.as_str(),
                    "__j_println" | "__j_print" | "__java_println" | "__java_print"
                ) {
                    for arg in &mut *args {
                        if java_print_arg_needs_tostring(
                            &arg.value,
                            tostring_classes,
                            enum_values,
                            current_class,
                            locals,
                        ) {
                            let receiver = arg.value.clone();
                            arg.value = java_tostring_call(receiver);
                        }
                    }
                }
                if fn_name == "__java_string_concat" {
                    for arg in &mut *args {
                        if java_expr_enum_type(&arg.value, enum_values, current_class, locals)
                            .is_some()
                        {
                            let receiver = arg.value.clone();
                            arg.value = java_tostring_call(receiver);
                        }
                    }
                }
            }
            if let ExprKind::Member { object, field, .. } = &mut callee.kind {
                if args.is_empty() {
                    if let ExprKind::Ident(ref name) = object.kind {
                        if java_record_has_component(locals.get(name).map(String::as_str), field) {
                            *expr = Expression::new(ExprKind::Member {
                                object: Box::new((**object).clone()),
                                field: java_record_storage_field(field),
                                null_safe: false,
                            });
                            return;
                        }
                    }
                }
                if field == "clear" && args.is_empty() {
                    if let ExprKind::Call {
                        callee: collection_callee,
                        args: collection_args,
                        ..
                    } = &object.kind
                    {
                        if collection_args.is_empty() {
                            if let ExprKind::Member {
                                object: map_object,
                                field: collection_field,
                                ..
                            } = &collection_callee.kind
                            {
                                if matches!(
                                    collection_field.as_str(),
                                    "keySet" | "values" | "entrySet"
                                ) {
                                    if let ExprKind::Ident(ref name) = map_object.kind {
                                        if java_type_is_map(locals.get(name).map(String::as_str)) {
                                            *expr = Expression::new(ExprKind::Call {
                                                callee: Box::new(Expression::ident(
                                                    "__java_map_clear",
                                                )),
                                                args: vec![Argument::positional(
                                                    (**map_object).clone(),
                                                )],
                                                optional: false,
                                            });
                                            return;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                if field == "remove" && args.len() == 1 {
                    if let ExprKind::Call {
                        callee: key_set_callee,
                        args: key_set_args,
                        ..
                    } = &object.kind
                    {
                        if key_set_args.is_empty() {
                            if let ExprKind::Member {
                                object: map_object,
                                field: key_set_field,
                                ..
                            } = &key_set_callee.kind
                            {
                                if key_set_field == "keySet" {
                                    if let ExprKind::Ident(ref name) = map_object.kind {
                                        if java_type_is_map(locals.get(name).map(String::as_str)) {
                                            *expr = Expression::new(ExprKind::Call {
                                                callee: Box::new(Expression::ident(
                                                    "__java_map_key_set_remove",
                                                )),
                                                args: vec![
                                                    Argument::positional((**map_object).clone()),
                                                    args[0].clone(),
                                                ],
                                                optional: false,
                                            });
                                            return;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                rewrite_java_tostring_expr(
                    object,
                    tostring_classes,
                    enum_values,
                    current_class,
                    locals,
                );
                if matches!(field.as_str(), "hasMoreElements" | "hasMoreTokens") {
                    if let ExprKind::Call { callee, .. } = &object.kind {
                        if matches!(
                            &callee.kind,
                            ExprKind::Ident(n)
                                if matches!(n.as_str(), "__j_props_keys" | "__j_props_elements" | "__j_props_enum")
                        ) {
                            *expr = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::ident("__j_enum_has_more")),
                                args: vec![Argument::positional((**object).clone())],
                                optional: false,
                            });
                            return;
                        }
                    }
                }
                if let ExprKind::Ident(ref name) = object.kind {
                    if java_type_is_enum_set(locals.get(name).map(String::as_str)) {
                        if let Some(internal) = java_enum_set_method_name(field) {
                            let mut new_args = Vec::with_capacity(args.len() + 1);
                            new_args.push(Argument::positional((**object).clone()));
                            new_args.extend(args.iter().cloned().map(|mut arg| {
                                if matches!(field.as_str(), "add" | "contains" | "remove") {
                                    if let Some(name_expr) =
                                        java_enum_member_arg_to_name(&arg.value, enum_values)
                                    {
                                        arg.value = name_expr;
                                    }
                                }
                                arg
                            }));
                            *expr = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::ident(internal)),
                                args: new_args,
                                optional: false,
                            });
                            return;
                        }
                    }
                    if java_type_is_uuid(locals.get(name).map(String::as_str)) {
                        if let Some(internal) = java_uuid_method_name(field) {
                            let mut new_args = Vec::with_capacity(args.len() + 1);
                            new_args.push(Argument::positional((**object).clone()));
                            new_args.extend(args.iter().cloned());
                            *expr = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::ident(internal)),
                                args: new_args,
                                optional: false,
                            });
                            return;
                        }
                    }
                    if java_type_is_instant(locals.get(name).map(String::as_str)) {
                        if let Some(internal) = java_instant_method_name(field) {
                            let mut new_args = Vec::with_capacity(args.len() + 1);
                            new_args.push(Argument::positional((**object).clone()));
                            new_args.extend(args.iter().cloned());
                            *expr = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::ident(internal)),
                                args: new_args,
                                optional: false,
                            });
                            return;
                        }
                    }
                    if java_type_is_time_value(locals.get(name).map(String::as_str)) {
                        if let Some(internal) = java_time_method_name(field) {
                            let mut new_args = Vec::with_capacity(args.len() + 1);
                            new_args.push(Argument::positional((**object).clone()));
                            new_args.extend(args.iter().cloned());
                            *expr = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::ident(internal)),
                                args: new_args,
                                optional: false,
                            });
                            return;
                        }
                    }
                    if java_type_is_duration(locals.get(name).map(String::as_str)) {
                        if let Some(internal) = java_duration_method_name(field) {
                            let mut new_args = Vec::with_capacity(args.len() + 1);
                            new_args.push(Argument::positional((**object).clone()));
                            new_args.extend(args.iter().cloned());
                            *expr = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::ident(internal)),
                                args: new_args,
                                optional: false,
                            });
                            return;
                        }
                    }
                    if java_type_is_zone_id(locals.get(name).map(String::as_str)) {
                        if let Some(internal) = java_zone_method_name(field) {
                            let mut new_args = Vec::with_capacity(args.len() + 1);
                            new_args.push(Argument::positional((**object).clone()));
                            new_args.extend(args.iter().cloned());
                            *expr = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::ident(internal)),
                                args: new_args,
                                optional: false,
                            });
                            return;
                        }
                    }
                    if java_type_is_bitset(locals.get(name).map(String::as_str)) {
                        if let Some(internal) = java_bitset_method_name(field) {
                            let mut new_args = Vec::with_capacity(args.len() + 1);
                            new_args.push(Argument::positional((**object).clone()));
                            new_args.extend(args.iter().cloned());
                            *expr = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::ident(internal)),
                                args: new_args,
                                optional: false,
                            });
                            return;
                        }
                    }
                    if java_type_is_list_like(locals.get(name).map(String::as_str)) {
                        if let Some(internal) = java_list_method_name(field, args.len()) {
                            let mut new_args = Vec::with_capacity(args.len() + 1);
                            let receiver = java_static_field_receiver(object, current_class)
                                .unwrap_or_else(|| (**object).clone());
                            new_args.push(Argument::positional(receiver));
                            new_args.extend(args.iter().cloned());
                            *expr = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::ident(internal)),
                                args: new_args,
                                optional: false,
                            });
                            return;
                        }
                    }
                    if java_type_is_map(locals.get(name).map(String::as_str)) {
                        if java_type_simple_name(
                            locals.get(name).map(String::as_str).unwrap_or_default(),
                        ) == "Properties"
                        {
                            if let Some(internal) = java_properties_method_name(field) {
                                let mut new_args = Vec::with_capacity(args.len() + 1);
                                new_args.push(Argument::positional((**object).clone()));
                                new_args.extend(args.iter().cloned());
                                *expr = Expression::new(ExprKind::Call {
                                    callee: Box::new(Expression::ident(internal)),
                                    args: new_args,
                                    optional: false,
                                });
                                return;
                            }
                        }
                        if field == "keySet"
                            && matches!(
                                locals.get(name).map(String::as_str),
                                Some("TreeMap") | Some("java.util.TreeMap")
                            )
                        {
                            *expr = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::ident("__java_sorted_map_key_set")),
                                args: vec![Argument::positional((**object).clone())],
                                optional: false,
                            });
                            return;
                        }
                        if java_type_is_hashtable(locals.get(name).map(String::as_str)) {
                            let internal = match field.as_str() {
                                "put" => Some("__java_hashtable_put"),
                                "keys" => Some("__java_hashtable_keys"),
                                "elements" => Some("__java_hashtable_elements"),
                                _ => None,
                            };
                            if let Some(internal) = internal {
                                let mut new_args = Vec::with_capacity(args.len() + 1);
                                new_args.push(Argument::positional((**object).clone()));
                                new_args.extend(args.iter().cloned());
                                *expr = Expression::new(ExprKind::Call {
                                    callee: Box::new(Expression::ident(internal)),
                                    args: new_args,
                                    optional: false,
                                });
                                return;
                            }
                        }
                        if let Some(internal) = java_map_method_name(field) {
                            let mut new_args = Vec::with_capacity(args.len() + 1);
                            let receiver = java_static_field_receiver(object, current_class)
                                .unwrap_or_else(|| (**object).clone());
                            new_args.push(Argument::positional(receiver));
                            new_args.extend(args.iter().cloned());
                            *expr = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::ident(internal)),
                                args: new_args,
                                optional: false,
                            });
                            return;
                        }
                    }
                    if java_type_is_priority_queue(locals.get(name).map(String::as_str)) {
                        let internal = match field.as_str() {
                            "add" | "offer" => Some("__java_priority_add"),
                            "poll" => Some("__java_sorted_poll"),
                            "peek" | "element" => Some("__java_priority_peek"),
                            "remove" if args.len() == 1 => Some("__java_set_remove"),
                            _ => None,
                        };
                        if let Some(internal) = internal {
                            let mut new_args = Vec::with_capacity(args.len() + 1);
                            new_args.push(Argument::positional((**object).clone()));
                            new_args.extend(args.iter().cloned());
                            *expr = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::ident(internal)),
                                args: new_args,
                                optional: false,
                            });
                            return;
                        }
                    }
                    if java_type_is_stack(locals.get(name).map(String::as_str)) {
                        let internal = match field.as_str() {
                            "push" => Some("__java_stack_push"),
                            "pop" => Some("__java_stack_pop"),
                            "peek" => Some("__java_stack_peek"),
                            "elementAt" => Some("__java_stack_element_at"),
                            "firstElement" => Some("__java_stack_first_element"),
                            "lastElement" => Some("__java_stack_last_element"),
                            "search" => Some("__java_stack_search"),
                            "set" => Some("__java_stack_set"),
                            "clone" => Some("__java_stack_clone"),
                            "remove" if args.len() == 1 => Some("__java_list_remove"),
                            _ => None,
                        };
                        if let Some(internal) = internal {
                            let mut new_args = Vec::with_capacity(args.len() + 1);
                            new_args.push(Argument::positional((**object).clone()));
                            new_args.extend(args.iter().cloned());
                            *expr = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::ident(internal)),
                                args: new_args,
                                optional: false,
                            });
                            return;
                        }
                    }
                    if java_type_is_vector(locals.get(name).map(String::as_str)) {
                        let internal = match field.as_str() {
                            "elementAt" => Some("__java_stack_element_at"),
                            "firstElement" => Some("__java_stack_first_element"),
                            "lastElement" => Some("__java_stack_last_element"),
                            "clone" => Some("__java_stack_clone"),
                            "indexOf" => Some("__java_list_index_of"),
                            "capacity" => Some("__java_vector_capacity"),
                            "ensureCapacity" => Some("__java_vector_ensure_capacity"),
                            "trimToSize" => Some("__java_vector_trim_to_size"),
                            "setSize" => Some("__java_vector_set_size"),
                            "elements" => Some("__java_vector_elements"),
                            _ => None,
                        };
                        if let Some(internal) = internal {
                            let mut new_args = Vec::with_capacity(args.len() + 1);
                            new_args.push(Argument::positional((**object).clone()));
                            new_args.extend(args.iter().cloned());
                            *expr = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::ident(internal)),
                                args: new_args,
                                optional: false,
                            });
                            return;
                        }
                    }
                    if java_type_is_enumeration(locals.get(name).map(String::as_str)) {
                        let internal = match field.as_str() {
                            "hasMoreElements" | "hasMoreTokens" => {
                                Some("__java_enumeration_has_more")
                            }
                            "nextElement" | "nextToken" => Some("__java_enumeration_next"),
                            _ => None,
                        };
                        if let Some(internal) = internal {
                            let mut new_args = Vec::with_capacity(args.len() + 1);
                            new_args.push(Argument::positional((**object).clone()));
                            new_args.extend(args.iter().cloned());
                            *expr = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::ident(internal)),
                                args: new_args,
                                optional: false,
                            });
                            return;
                        }
                    }
                    if java_type_is_queue_or_deque(locals.get(name).map(String::as_str)) {
                        let internal = java_queue_method_name(
                            locals.get(name).map(String::as_str),
                            field,
                            args.len(),
                        );
                        if let Some(internal) = internal {
                            let mut new_args = Vec::with_capacity(args.len() + 1);
                            new_args.push(Argument::positional((**object).clone()));
                            new_args.extend(args.iter().cloned());
                            *expr = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::ident(internal)),
                                args: new_args,
                                optional: false,
                            });
                            return;
                        }
                    }
                    if java_type_is_set(locals.get(name).map(String::as_str)) {
                        let simple_type = locals
                            .get(name)
                            .map(String::as_str)
                            .map(java_type_simple_name)
                            .unwrap_or_default();
                        let internal = match field.as_str() {
                            "add" if matches!(simple_type, "TreeSet") => Some("__java_sorted_add"),
                            "add" => Some("__java_set_add"),
                            "remove" => Some("__java_set_remove"),
                            _ => None,
                        };
                        if let Some(internal) = internal {
                            let mut new_args = Vec::with_capacity(args.len() + 1);
                            new_args.push(Argument::positional((**object).clone()));
                            new_args.extend(args.iter().cloned());
                            *expr = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::ident(internal)),
                                args: new_args,
                                optional: false,
                            });
                            return;
                        }
                    }
                }
                if matches!(field.as_str(), "add" | "offer" | "poll")
                    && matches!(
                        object.kind,
                        ExprKind::Ident(ref name)
                            if matches!(
                                locals.get(name).map(String::as_str),
                                Some("TreeSet")
                                    | Some("java.util.TreeSet")
                            )
                    )
                {
                    let internal = if field == "poll" {
                        "__java_sorted_poll"
                    } else {
                        "__java_sorted_add"
                    };
                    let mut new_args = Vec::with_capacity(args.len() + 1);
                    new_args.push(Argument::positional((**object).clone()));
                    new_args.extend(args.iter().cloned());
                    *expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident(internal)),
                        args: new_args,
                        optional: false,
                    });
                    return;
                }
                if let Some(name) = java_bigint_method_name(field) {
                    if java_expr_is_bigint(object, locals) {
                        let mut new_args = Vec::with_capacity(args.len() + 1);
                        new_args.push(Argument::positional((**object).clone()));
                        new_args.extend(args.iter().cloned());
                        *expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident(name)),
                            args: new_args,
                            optional: false,
                        });
                        return;
                    }
                }
                if let Some(name) = java_bigdecimal_method_name(field, args.len()) {
                    if java_expr_is_bigdecimal(object, locals) {
                        let mut new_args = Vec::with_capacity(args.len() + 1);
                        new_args.push(Argument::positional((**object).clone()));
                        new_args.extend(args.iter().cloned());
                        *expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident(name)),
                            args: new_args,
                            optional: false,
                        });
                        return;
                    }
                }
                if field == "toString"
                    && java_expr_has_user_tostring(object, tostring_classes, current_class, locals)
                {
                    *field = "tostring".to_string();
                }
                if field == "toString"
                    && java_expr_enum_type(object, enum_values, current_class, locals).is_some()
                {
                    *field = "tostring".to_string();
                }
            } else {
                rewrite_java_tostring_expr(
                    callee,
                    tostring_classes,
                    enum_values,
                    current_class,
                    locals,
                );
            }
        }
        ExprKind::Member { object, field, .. } => {
            if java_member_chain_ends_with(object, "Boolean") {
                if field == "TRUE" {
                    *expr = Expression::bool(true);
                    return;
                }
                if field == "FALSE" {
                    *expr = Expression::bool(false);
                    return;
                }
            }
            // java.util.regex.Pattern flag constants (JLS values).
            if java_member_chain_ends_with(object, "Pattern") {
                let flag = match field.as_str() {
                    "UNIX_LINES" => Some(1),
                    "CASE_INSENSITIVE" => Some(2),
                    "COMMENTS" => Some(4),
                    "MULTILINE" => Some(8),
                    "LITERAL" => Some(16),
                    "DOTALL" => Some(32),
                    "UNICODE_CASE" => Some(64),
                    "CANON_EQ" => Some(128),
                    "UNICODE_CHARACTER_CLASS" => Some(256),
                    _ => None,
                };
                if let Some(value) = flag {
                    *expr = Expression::int(value);
                    return;
                }
            }
            if matches!(field.as_str(), "SECONDS" | "MILLIS" | "MINUTES" | "HOURS")
                && java_member_chain_ends_with(object, "ChronoUnit")
            {
                *expr = Expression::string(field);
                return;
            }
            if field == "UTC" && java_member_chain_ends_with(object, "ZoneOffset") {
                *expr = Expression::string("Z");
                return;
            }
            if field == "SHORT_IDS" && java_member_chain_ends_with(object, "ZoneId") {
                *expr = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__java_zone_id_short_ids")),
                    args: vec![],
                    optional: false,
                });
                return;
            }
            rewrite_java_tostring_expr(
                object,
                tostring_classes,
                enum_values,
                current_class,
                locals,
            );
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_java_tostring_expr(
                object,
                tostring_classes,
                enum_values,
                current_class,
                locals,
            );
            rewrite_java_tostring_expr(index, tostring_classes, enum_values, current_class, locals);
        }
        ExprKind::Binary { left, right, .. } => {
            rewrite_java_tostring_expr(left, tostring_classes, enum_values, current_class, locals);
            rewrite_java_tostring_expr(right, tostring_classes, enum_values, current_class, locals);
            rewrite_java_switch_enum_label(left, right, locals);
            rewrite_java_switch_enum_label(right, left, locals);
        }
        ExprKind::Unary { expr: inner, .. } => {
            rewrite_java_tostring_expr(inner, tostring_classes, enum_values, current_class, locals);
        }
        ExprKind::Lambda { params, body, .. } => {
            let mut lambda_locals = locals.clone();
            for param in params {
                if let Some(type_hint) = &param.type_hint {
                    lambda_locals.insert(param.name.clone(), type_hint.clone());
                }
            }
            match body {
                LambdaBody::Expr(inner) => rewrite_java_tostring_expr(
                    inner,
                    tostring_classes,
                    enum_values,
                    current_class,
                    &lambda_locals,
                ),
                LambdaBody::Block(stmts) => rewrite_java_tostring_stmts(
                    stmts,
                    tostring_classes,
                    enum_values,
                    current_class,
                    &mut lambda_locals,
                ),
            }
        }
        ExprKind::IsType {
            expr: inner,
            type_name,
        } => {
            rewrite_java_tostring_expr(inner, tostring_classes, enum_values, current_class, locals);
            if enum_values.contains_key(java_type_simple_name(type_name)) {
                if java_enum_type_from_member_expr(inner).map(java_type_simple_name)
                    == Some(java_type_simple_name(type_name))
                {
                    *expr = Expression::bool(true);
                }
            }
        }
        ExprKind::Assign { target, value } => {
            rewrite_java_tostring_expr(value, tostring_classes, enum_values, current_class, locals);
            rewrite_java_tostring_expr(
                target,
                tostring_classes,
                enum_values,
                current_class,
                locals,
            );
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_java_tostring_expr(cond, tostring_classes, enum_values, current_class, locals);
            rewrite_java_tostring_expr(then, tostring_classes, enum_values, current_class, locals);
            rewrite_java_tostring_expr(else_, tostring_classes, enum_values, current_class, locals);
        }
        ExprKind::Array(elems) => {
            for elem in elems {
                rewrite_java_tostring_expr(
                    &mut elem.value,
                    tostring_classes,
                    enum_values,
                    current_class,
                    locals,
                );
            }
        }
        ExprKind::New { args, .. } => {
            for arg in args {
                rewrite_java_tostring_expr(
                    &mut arg.value,
                    tostring_classes,
                    enum_values,
                    current_class,
                    locals,
                );
            }
            // Built-in java exceptions the shared emitter doesn't know:
            // build the canonical exception object with the JLS supertype
            // chain in __types (parents pre-canonicalized) so catch-clause
            // REF_TEST subtype matching works.
            if let ExprKind::New { class, args } = &expr.kind {
                if let ExprKind::Ident(class_name) = &class.kind {
                    let simple = java_type_simple_name(class_name);
                    // Reroute java-only types always; shared-known types only
                    // for the `(message, cause)` ctor the shared path drops.
                    let chain: Option<&'static [&'static str]> =
                        java_builtin_exception_chain(simple).or_else(|| {
                            if args.len() >= 2 {
                                java_known_exception_chain(simple)
                            } else {
                                None
                            }
                        });
                    if let Some(chain) = chain {
                        let mut chain_elems = vec![ArrayElement {
                            key: None,
                            value: Expression::string(simple),
                            spread: false,
                            by_ref: false,
                        }];
                        chain_elems.extend(chain.iter().map(|parent| ArrayElement {
                            key: None,
                            value: Expression::string(parent),
                            spread: false,
                            by_ref: false,
                        }));
                        let mut call_args = vec![
                            Argument::positional(Expression::string(simple)),
                            Argument::positional(Expression::new(ExprKind::Array(chain_elems))),
                        ];
                        call_args.extend(args.iter().cloned());
                        *expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident("__j_exc")),
                            args: call_args,
                            optional: false,
                        });
                    }
                }
            }
        }
        _ => {}
    }
}

/// JLS supertype chains for built-in exceptions the shared emitter's
/// `is_exception_type` doesn't recognize. Parents use the CANONICAL names
/// the catch-clause compiler compares against (`RuntimeException` →
/// "RuntimeError", `IOException` → "IOError").
fn java_builtin_exception_chain(name: &str) -> Option<&'static [&'static str]> {
    const RT: &[&str] = &["RuntimeError", "Exception", "Throwable"];
    const IOOB: &[&str] = &[
        "IndexOutOfBoundsException",
        "RuntimeError",
        "Exception",
        "Throwable",
    ];
    const CHECKED: &[&str] = &["Exception", "Throwable"];
    Some(match name {
        "IllegalArgumentException"
        | "IllegalStateException"
        | "ClassCastException"
        | "UnsupportedOperationException"
        | "NullPointerException"
        | "IndexOutOfBoundsException"
        | "ArithmeticException"
        | "NegativeArraySizeException"
        | "SecurityException"
        | "NoSuchElementException"
        | "ConcurrentModificationException"
        | "ArrayStoreException" => RT,
        "NumberFormatException" => &[
            "IllegalArgumentException",
            "RuntimeError",
            "Exception",
            "Throwable",
        ],
        "ArrayIndexOutOfBoundsException" | "StringIndexOutOfBoundsException" => IOOB,
        "InterruptedException"
        | "CloneNotSupportedException"
        | "NoSuchMethodException"
        | "NoSuchFieldException"
        | "ClassNotFoundException"
        | "ParseException" => CHECKED,
        _ => return None,
    })
}

/// Canonical chains for exception types the shared emitter DOES recognize
/// (their catch clauses compare canonical names).
fn java_known_exception_chain(name: &str) -> Option<&'static [&'static str]> {
    Some(match name {
        "RuntimeException" => &["RuntimeError", "Exception", "Throwable"],
        "Exception" => &["Exception", "Throwable"],
        "Throwable" => &["Throwable"],
        "IOException" => &["IOError", "Exception", "Throwable"],
        "FileNotFoundException" => &["FileNotFoundError", "IOError", "Exception", "Throwable"],
        _ => return None,
    })
}

/// Full supertype list (excluding self) for a built-in exception name,
/// canonicalized where the shared emitter has a canonical form.
fn java_exception_supertypes(name: &str) -> Option<Vec<String>> {
    if let Some(chain) = java_known_exception_chain(name) {
        return Some(chain.iter().map(|s| s.to_string()).collect());
    }
    java_builtin_exception_chain(name).map(|chain| {
        std::iter::once(name.to_string())
            .chain(chain.iter().map(|s| s.to_string()))
            .collect()
    })
}

fn rewrite_java_switch_enum_label(
    maybe_switch_value: &Expression,
    maybe_label: &mut Expression,
    locals: &std::collections::HashMap<String, String>,
) {
    let switch_type = match &maybe_switch_value.kind {
        ExprKind::Ident(name) if name.starts_with("__java_switch_value_") => locals.get(name),
        _ => None,
    };
    let Some(type_hint) = switch_type else {
        return;
    };
    let ExprKind::Ident(label) = &maybe_label.kind else {
        return;
    };
    if label.starts_with("__") {
        return;
    }
    *maybe_label = Expression::new(ExprKind::Member {
        object: Box::new(Expression::ident(java_type_simple_name(type_hint))),
        field: label.clone(),
        null_safe: false,
    });
}

fn rewrite_java_enum_set_static_call(
    callee: &Expression,
    args: &[Argument],
    enum_values: &HashMap<String, Vec<String>>,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    if !java_member_chain_ends_with(object, "EnumSet") {
        return None;
    }
    let names = match field.as_str() {
        "noneOf" | "allOf" => {
            let enum_name = args
                .first()
                .and_then(|arg| java_string_literal(&arg.value))?;
            java_enum_names_expr(enum_values, enum_name)?
        }
        "of" => {
            let enum_name = args
                .first()
                .and_then(|arg| java_enum_type_from_member_expr(&arg.value))?;
            java_enum_names_expr(enum_values, enum_name)?
        }
        "range" => {
            let enum_name = args
                .first()
                .and_then(|arg| java_enum_type_from_member_expr(&arg.value))?;
            java_enum_names_expr(enum_values, enum_name)?
        }
        "copyOf" | "complementOf" => Expression::null(),
        _ => return None,
    };
    let internal = match field.as_str() {
        "noneOf" => "__java_enum_set_none_of",
        "allOf" => "__java_enum_set_all_of",
        "of" => "__java_enum_set_of",
        "copyOf" => "__java_enum_set_copy_of",
        "complementOf" => "__java_enum_set_complement_of",
        "range" => "__java_enum_set_range",
        _ => return None,
    };
    let mut new_args = Vec::with_capacity(args.len() + 1);
    match field.as_str() {
        "copyOf" | "complementOf" => new_args.extend(args.iter().cloned()),
        "of" | "range" => {
            new_args.push(Argument::positional(names));
            new_args.extend(args.iter().cloned().map(|mut arg| {
                if let Some(name_expr) = java_enum_member_arg_to_name(&arg.value, enum_values) {
                    arg.value = name_expr;
                }
                arg
            }));
        }
        _ => new_args.push(Argument::positional(names)),
    }
    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(internal)),
        args: new_args,
        optional: false,
    }))
}

fn java_member_chain_ends_with(expr: &Expression, expected: &str) -> bool {
    let mut parts = Vec::new();
    collect_member_chain(expr, &mut parts).is_some() && parts.last().copied() == Some(expected)
}

fn java_string_literal(expr: &Expression) -> Option<&str> {
    if let ExprKind::Lit(Literal::Str(value)) = &expr.kind {
        Some(value.as_str())
    } else {
        None
    }
}

fn java_enum_type_from_member_expr(expr: &Expression) -> Option<&str> {
    if let ExprKind::Member { object, .. } = &expr.kind {
        if let ExprKind::Ident(name) = &object.kind {
            return Some(name.as_str());
        }
    }
    None
}

/// EnumSet internals take ordinal indices at value boundaries
/// (`names[value]` lookups). Convert an enum-constant expression
/// (`Color.GREEN`) to its declaration index, and an enum-typed variable
/// to an `.ordinal()` call.
fn java_enum_member_arg_to_name(
    arg: &Expression,
    enum_values: &HashMap<String, Vec<String>>,
) -> Option<Expression> {
    if let ExprKind::Member { object, field, .. } = &arg.kind {
        if let ExprKind::Ident(type_name) = &object.kind {
            let base = type_name.rsplit('.').next().unwrap_or(type_name);
            if let Some(index) = enum_values
                .get(base)
                .and_then(|members| members.iter().position(|m| m == field))
            {
                return Some(Expression::int(index as i64));
            }
        }
    }
    None
}

fn java_enum_names_expr(
    enum_values: &HashMap<String, Vec<String>>,
    enum_name: &str,
) -> Option<Expression> {
    let base = enum_name.rsplit('.').next().unwrap_or(enum_name);
    let values = enum_values.get(base)?;
    Some(Expression::new(ExprKind::Array(
        values
            .iter()
            .map(|name| ArrayElement {
                key: None,
                value: Expression::string(name),
                spread: false,
                by_ref: false,
            })
            .collect(),
    )))
}

fn java_type_is_enum_set(type_name: Option<&str>) -> bool {
    let Some(type_name) = type_name else {
        return false;
    };
    let base = type_name
        .rsplit('.')
        .next()
        .unwrap_or(type_name)
        .split('<')
        .next()
        .unwrap_or(type_name)
        .trim();
    base == "EnumSet"
}

fn java_enum_set_method_name(method: &str) -> Option<&'static str> {
    Some(match method {
        "add" => "__java_enum_set_add",
        "addAll" => "__java_enum_set_add_all",
        "contains" => "__java_enum_set_contains",
        "containsAll" => "__java_enum_set_contains_all",
        "remove" => "__java_enum_set_remove",
        "equals" => "__java_enum_set_equals",
        "hashCode" => "__java_enum_set_hash_code",
        "iterator" => "__java_enum_set_iterator",
        "getClass" => "__java_enum_set_get_class",
        _ => return None,
    })
}

fn java_type_is_map(type_name: Option<&str>) -> bool {
    let Some(type_name) = type_name else {
        return false;
    };
    let base = type_name
        .rsplit('.')
        .next()
        .unwrap_or(type_name)
        .split('<')
        .next()
        .unwrap_or(type_name)
        .trim();
    matches!(
        base,
        "Map"
            | "HashMap"
            | "IdentityHashMap"
            | "LinkedHashMap"
            | "ConcurrentHashMap"
            | "WeakHashMap"
            | "TreeMap"
            | "SortedMap"
            | "NavigableMap"
            | "Hashtable"
            | "Properties"
    )
}

fn java_type_is_set(type_name: Option<&str>) -> bool {
    let Some(type_name) = type_name else {
        return false;
    };
    let base = type_name
        .rsplit('.')
        .next()
        .unwrap_or(type_name)
        .split('<')
        .next()
        .unwrap_or(type_name)
        .trim();
    matches!(
        base,
        "Set" | "HashSet" | "LinkedHashSet" | "TreeSet" | "SortedSet" | "NavigableSet"
    )
}

fn java_type_is_priority_queue(type_name: Option<&str>) -> bool {
    let Some(type_name) = type_name else {
        return false;
    };
    java_type_simple_name(type_name) == "PriorityQueue"
}

fn java_type_is_stack(type_name: Option<&str>) -> bool {
    let Some(type_name) = type_name else {
        return false;
    };
    java_type_simple_name(type_name) == "Stack"
}

fn java_type_is_vector(type_name: Option<&str>) -> bool {
    let Some(type_name) = type_name else {
        return false;
    };
    java_type_simple_name(type_name) == "Vector"
}

fn java_type_is_enumeration(type_name: Option<&str>) -> bool {
    let Some(type_name) = type_name else {
        return false;
    };
    java_type_simple_name(type_name) == "Enumeration"
}

fn java_type_is_hashtable(type_name: Option<&str>) -> bool {
    let Some(type_name) = type_name else {
        return false;
    };
    java_type_simple_name(type_name) == "Hashtable"
}

fn java_type_is_list_like(type_name: Option<&str>) -> bool {
    let Some(type_name) = type_name else {
        return false;
    };
    let simple = java_type_simple_name(type_name);
    let simple = simple.split('<').next().unwrap_or(simple);
    if !type_name.contains('.') && java_user_defined_type(simple) {
        return false;
    }
    matches!(
        simple,
        "List" | "ArrayList" | "LinkedList" | "CopyOnWriteArrayList" | "Vector"
    )
}

fn java_user_defined_type(simple: &str) -> bool {
    JAVA_REFLECTION_CLASSES.with(|classes| classes.borrow().contains_key(simple))
        || JAVA_INTERFACE_NAMES.with(|names| names.borrow().contains(simple))
        || JAVA_ENUM_VALUES.with(|values| values.borrow().contains_key(simple))
}

fn java_list_method_name(method: &str, argc: usize) -> Option<&'static str> {
    Some(match method {
        "add" if argc == 1 || argc == 2 => "__java_list_add",
        "addIfAbsent" if argc == 1 => "__java_copy_on_write_add_if_absent",
        "size" if argc == 0 => "__java_list_size",
        "remove" if argc == 1 => "__java_list_remove",
        "containsAll" if argc == 1 => "__java_list_contains_all",
        "equals" if argc == 1 => "__java_list_equals",
        "indexOf" if argc == 1 => "__java_list_index_of",
        "spliterator" if argc == 0 => "__j_spliterator_from_array",
        _ => return None,
    })
}

fn java_type_is_spliterator(type_name: Option<&str>) -> bool {
    let Some(type_name) = type_name else {
        return false;
    };
    let simple = java_type_simple_name(type_name);
    simple.split('<').next().unwrap_or(simple) == "Spliterator"
}

fn java_type_is_random(type_name: Option<&str>) -> bool {
    let Some(type_name) = type_name else {
        return false;
    };
    matches!(
        java_type_simple_name(type_name),
        "Random" | "SplittableRandom" | "ThreadLocalRandom"
    )
}

fn java_random_receiver(receiver: &Expression) -> bool {
    match &receiver.kind {
        ExprKind::Ident(name) => {
            JAVA_RANDOM_VARS.with(|vars| vars.borrow().contains(name.as_str()))
        }
        ExprKind::Call { callee, .. } => {
            matches!(&callee.kind, ExprKind::Ident(name) if name == "__java_random_new")
        }
        _ => false,
    }
}

fn java_list_result_receiver(expr: &Expression) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Call { callee, .. }
            if matches!(&callee.kind, ExprKind::Ident(name) if name == "__j_pb_command_get")
    )
}

fn java_map_result_receiver(expr: &Expression) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Call { callee, .. }
            if matches!(&callee.kind, ExprKind::Ident(name) if name == "__j_pb_environment")
    )
}

fn java_spliterator_receiver(expr: &Expression) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Call { callee, .. }
            if matches!(&callee.kind, ExprKind::Ident(name) if name == "__j_spliterator_from_array")
    )
}

fn java_spliterator_method_name(method: &str, argc: usize) -> Option<&'static str> {
    Some(match method {
        "estimateSize" if argc == 0 => "__j_spliterator_estimate_size",
        "tryAdvance" if argc == 1 => "__j_spliterator_try_advance",
        "forEachRemaining" if argc == 1 => "__j_spliterator_for_each_remaining",
        "characteristics" if argc == 0 => "__j_spliterator_characteristics",
        "trySplit" if argc == 0 => "__j_spliterator_try_split",
        "getComparator" if argc == 0 => "__j_spliterator_get_comparator",
        _ => return None,
    })
}

fn java_spliterator_constant(name: &str) -> Option<i64> {
    Some(match name {
        "DISTINCT" => 0x0001,
        "SORTED" => 0x0004,
        "ORDERED" => 0x0010,
        "SIZED" => 0x0040,
        "NONNULL" => 0x0100,
        "IMMUTABLE" => 0x0400,
        "CONCURRENT" => 0x1000,
        "SUBSIZED" => 0x4000,
        _ => return None,
    })
}

fn java_type_is_runtime(type_name: Option<&str>) -> bool {
    type_name
        .map(|type_name| {
            let simple = java_type_simple_name(type_name);
            simple.split('<').next().unwrap_or(simple) == "Runtime"
        })
        .unwrap_or(false)
}

fn java_runtime_receiver(expr: &Expression) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Call { callee, .. }
            if matches!(&callee.kind, ExprKind::Ident(name) if name == "__j_runtime_get")
    )
}

fn java_runtime_method_name(method: &str, argc: usize) -> Option<&'static str> {
    Some(match method {
        "availableProcessors" if argc == 0 => "__j_runtime_processors",
        "freeMemory" if argc == 0 => "__j_runtime_free_memory",
        "totalMemory" if argc == 0 => "__j_runtime_total_memory",
        "maxMemory" if argc == 0 => "__j_runtime_max_memory",
        "gc" if argc == 0 => "__j_runtime_noop",
        "runFinalization" if argc == 0 => "__j_runtime_noop",
        _ => return None,
    })
}

fn java_type_is_process_builder(type_name: Option<&str>) -> bool {
    type_name
        .map(|type_name| {
            let simple = java_type_simple_name(type_name);
            simple.split('<').next().unwrap_or(simple) == "ProcessBuilder"
        })
        .unwrap_or(false)
}

fn java_process_builder_receiver(expr: &Expression) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Call { callee, .. }
            if matches!(&callee.kind, ExprKind::Ident(name) if name == "__j_pb_new")
    )
}

fn java_process_builder_method_name(method: &str, argc: usize) -> Option<&'static str> {
    Some(match method {
        "command" if argc == 0 => "__j_pb_command_get",
        "command" if argc >= 1 => "__j_pb_command_set",
        "directory" if argc == 0 => "__j_pb_directory_get",
        "directory" if argc == 1 => "__j_pb_directory_set",
        "redirectErrorStream" if argc == 0 => "__j_pb_redirect_error_stream_get",
        "redirectErrorStream" if argc == 1 => "__j_pb_redirect_error_stream_set",
        "environment" if argc == 0 => "__j_pb_environment",
        "inheritIO" if argc == 0 => "__j_pb_inherit_io",
        "redirectInput" if argc == 0 => "__j_pb_redirect_input_get",
        "redirectOutput" if argc == 0 => "__j_pb_redirect_output_get",
        "redirectOutput" if argc == 1 => "__j_pb_redirect_output_set",
        "redirectError" if argc == 0 => "__j_pb_redirect_error_get",
        "start" if argc == 0 => "__j_pb_start",
        _ => return None,
    })
}

fn java_type_is_process(type_name: Option<&str>) -> bool {
    type_name
        .map(|type_name| {
            let simple = java_type_simple_name(type_name);
            simple.split('<').next().unwrap_or(simple) == "Process"
        })
        .unwrap_or(false)
}

fn java_process_receiver(expr: &Expression) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Call { callee, .. }
            if matches!(&callee.kind, ExprKind::Ident(name) if name == "__j_pb_start")
    )
}

fn java_process_method_name(method: &str, argc: usize) -> Option<&'static str> {
    Some(match method {
        "waitFor" if argc == 0 => "__j_process_wait_for",
        "isAlive" if argc == 0 => "__j_process_is_alive",
        "getInputStream" if argc == 0 => "__j_process_input_stream",
        "exitValue" if argc == 0 => "__j_process_exit_value",
        "destroy" if argc == 0 => "__j_process_destroy",
        _ => return None,
    })
}

fn java_type_is_file(type_name: Option<&str>) -> bool {
    type_name
        .map(|type_name| {
            let simple = java_type_simple_name(type_name);
            simple.split('<').next().unwrap_or(simple) == "File"
        })
        .unwrap_or(false)
}

fn java_file_receiver(expr: &Expression) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Call { callee, .. }
            if matches!(&callee.kind, ExprKind::Ident(name) if matches!(name.as_str(), "__j_file_new" | "__j_pb_directory_get"))
    )
}

fn java_type_is_redirect(type_name: Option<&str>) -> bool {
    type_name
        .map(|type_name| {
            let simple = java_type_simple_name(type_name);
            simple.split('<').next().unwrap_or(simple) == "Redirect"
        })
        .unwrap_or(false)
}

fn java_redirect_receiver(expr: &Expression) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Call { callee, .. }
            if matches!(&callee.kind, ExprKind::Ident(name) if matches!(
                name.as_str(),
                "__j_pb_redirect"
                    | "__j_pb_redirect_input_get"
                    | "__j_pb_redirect_output_get"
                    | "__j_pb_redirect_error_get"
            ))
    )
}

fn java_type_is_queue_or_deque(type_name: Option<&str>) -> bool {
    let Some(type_name) = type_name else {
        return false;
    };
    matches!(
        java_type_simple_name(type_name),
        "Queue" | "Deque" | "ArrayDeque" | "LinkedList" | "LinkedBlockingQueue"
    )
}

fn java_type_is_linked_blocking_queue(type_name: Option<&str>) -> bool {
    let Some(type_name) = type_name else {
        return false;
    };
    java_type_simple_name(type_name) == "LinkedBlockingQueue"
}

fn java_queue_method_name(
    type_name: Option<&str>,
    method: &str,
    argc: usize,
) -> Option<&'static str> {
    let is_blocking_queue = java_type_is_linked_blocking_queue(type_name);
    Some(if is_blocking_queue {
        match method {
            "add" if argc == 1 => "__java_blocking_queue_add",
            "offer" if argc == 1 || argc == 3 => "__java_blocking_queue_offer",
            "put" if argc == 1 => "__java_blocking_queue_put",
            "take" if argc == 0 => "__java_blocking_queue_take",
            "poll" if argc == 0 || argc == 2 => "__java_blocking_queue_poll",
            "remove" if argc == 0 => "__java_queue_remove_checked",
            "peek" if argc == 0 => "__java_queue_peek",
            "element" if argc == 0 => "__java_queue_element_checked",
            "remove" if argc == 1 => "__java_set_remove",
            _ => return None,
        }
    } else {
        match method {
            "add" | "offer" if argc == 1 => "__java_queue_add",
            "poll" if argc == 0 => "__java_queue_poll",
            "remove" if argc == 0 => "__java_queue_remove_checked",
            "peek" if argc == 0 => "__java_queue_peek",
            "element" if argc == 0 => "__java_queue_element_checked",
            "getFirst" | "peekFirst" if argc == 0 => "__java_stack_first_element",
            "getLast" | "peekLast" if argc == 0 => "__java_stack_last_element",
            "push" if argc == 1 => "__java_deque_push",
            "pop" if argc == 0 => "__java_deque_pop",
            "remove" if argc == 1 => "__java_set_remove",
            _ => return None,
        }
    })
}

fn java_type_is_semaphore(type_name: Option<&str>) -> bool {
    let Some(type_name) = type_name else {
        return false;
    };
    java_type_base_simple_name(type_name) == "Semaphore"
}

fn java_semaphore_method_name(method: &str) -> Option<&'static str> {
    Some(match method {
        "availablePermits" => "__java_semaphore_available",
        "acquire" | "acquireUninterruptibly" => "__java_semaphore_acquire",
        "release" => "__java_semaphore_release",
        "tryAcquire" => "__java_semaphore_try_acquire",
        "drainPermits" => "__java_semaphore_drain",
        "hasQueuedThreads" => "__java_semaphore_has_queued",
        "getQueueLength" => "__java_semaphore_queue_length",
        "isFair" => "__java_semaphore_is_fair",
        _ => return None,
    })
}

fn java_type_is_count_down_latch(type_name: Option<&str>) -> bool {
    let Some(type_name) = type_name else {
        return false;
    };
    java_type_base_simple_name(type_name) == "CountDownLatch"
}

fn java_count_down_latch_method_name(method: &str, argc: usize) -> Option<&'static str> {
    Some(match method {
        "countDown" if argc == 0 => "__j_latch_count_down",
        "getCount" if argc == 0 => "__j_latch_get_count",
        "await" if argc == 0 => "__j_latch_await",
        "await" if argc == 2 => "__j_latch_await_timeout",
        _ => return None,
    })
}

fn java_type_is_future_task(type_name: Option<&str>) -> bool {
    let Some(type_name) = type_name else {
        return false;
    };
    java_type_base_simple_name(type_name) == "FutureTask"
}

fn java_future_task_method_name(method: &str, argc: usize) -> Option<&'static str> {
    Some(match method {
        "run" if argc == 0 => "__j_future_task_run",
        "get" if argc == 0 || argc == 2 => "__j_future_task_get",
        "cancel" if argc == 1 => "__j_future_task_cancel",
        "isDone" if argc == 0 => "__j_future_task_is_done",
        "isCancelled" if argc == 0 => "__j_future_task_is_cancelled",
        "runAndReset" if argc == 0 => "__j_future_task_run_and_reset",
        _ => return None,
    })
}

fn java_type_is_executor_service(type_name: Option<&str>) -> bool {
    let Some(type_name) = type_name else {
        return false;
    };
    matches!(
        java_type_base_simple_name(type_name),
        "ExecutorService" | "Executor"
    )
}

fn java_executor_method_name(method: &str, argc: usize) -> Option<&'static str> {
    Some(match method {
        "submit" if argc == 1 => "__j_exec_submit",
        "execute" if argc == 1 => "__j_exec_execute",
        "shutdown" if argc == 0 => "__j_exec_shutdown",
        "shutdownNow" if argc == 0 => "__j_exec_shutdown_now",
        "isShutdown" if argc == 0 => "__j_exec_is_shutdown",
        "awaitTermination" if argc == 2 => "__j_exec_await",
        _ => return None,
    })
}

fn java_properties_method_name(method: &str) -> Option<&'static str> {
    Some(match method {
        "getProperty" => "__j_props_get",
        "setProperty" => "__j_props_set",
        "stringPropertyNames" => "__j_props_names",
        "keys" => "__j_props_keys",
        "elements" => "__j_props_elements",
        "getClass" => "__j_props_class",
        _ => return None,
    })
}

fn java_map_method_name(method: &str) -> Option<&'static str> {
    Some(match method {
        "get" => "__java_map_get",
        "put" => "__java_map_put",
        "putAll" => "__java_map_put_all",
        "remove" => "__java_map_remove",
        "getOrDefault" => "__java_map_get_or_default",
        "containsKey" => "__java_map_contains_key",
        "containsValue" => "__java_map_contains_value",
        "keySet" => "__java_map_key_set",
        "values" => "__java_map_values",
        "entrySet" => "__java_map_entry_set",
        "putIfAbsent" => "__java_map_put_if_absent",
        "computeIfAbsent" => "__java_map_compute_if_absent",
        "computeIfPresent" => "__java_map_compute_if_present",
        "compute" => "__java_map_compute",
        "merge" => "__java_map_merge",
        "replace" => "__java_map_replace",
        "replaceAll" => "__java_map_replace_all",
        "forEach" => "__java_map_for_each",
        "clear" => "__java_map_clear",
        "clone" => "__java_map_clone",
        "size" => "__java_map_size",
        "isEmpty" => "__java_map_is_empty",
        "equals" => "__java_map_equals",
        "firstKey" => "__java_sorted_map_first_key",
        "lastKey" => "__java_sorted_map_last_key",
        "firstEntry" => "__java_sorted_map_first_entry",
        "lastEntry" => "__java_sorted_map_last_entry",
        "ceilingEntry" => "__java_sorted_map_ceiling_entry",
        "floorEntry" => "__java_sorted_map_floor_entry",
        "higherEntry" => "__java_sorted_map_higher_entry",
        "lowerEntry" => "__java_sorted_map_lower_entry",
        "ceilingKey" => "__java_sorted_map_ceiling_key",
        "floorKey" => "__java_sorted_map_floor_key",
        "higherKey" => "__java_sorted_map_higher_key",
        "lowerKey" => "__java_sorted_map_lower_key",
        "pollFirstEntry" => "__java_sorted_map_poll_first_entry",
        "pollLastEntry" => "__java_sorted_map_poll_last_entry",
        "descendingKeySet" => "__java_sorted_map_descending_key_set",
        "descendingMap" => "__java_sorted_map_descending_map",
        "subMap" => "__java_map_sub_map",
        "headMap" => "__java_map_head_map",
        "tailMap" => "__java_map_tail_map",
        _ => return None,
    })
}

fn java_type_is_bitset(type_name: Option<&str>) -> bool {
    let Some(type_name) = type_name else {
        return false;
    };
    let base = type_name
        .rsplit('.')
        .next()
        .unwrap_or(type_name)
        .split('<')
        .next()
        .unwrap_or(type_name)
        .trim();
    base == "BitSet"
}

fn java_type_is_uuid(type_name: Option<&str>) -> bool {
    let Some(type_name) = type_name else {
        return false;
    };
    let base = type_name
        .rsplit('.')
        .next()
        .unwrap_or(type_name)
        .split('<')
        .next()
        .unwrap_or(type_name)
        .trim();
    base == "UUID"
}

fn java_type_is_instant(type_name: Option<&str>) -> bool {
    let Some(type_name) = type_name else {
        return false;
    };
    let base = type_name
        .rsplit('.')
        .next()
        .unwrap_or(type_name)
        .split('<')
        .next()
        .unwrap_or(type_name)
        .trim();
    base == "Instant"
}

fn java_type_is_time_value(type_name: Option<&str>) -> bool {
    let Some(type_name) = type_name else {
        return false;
    };
    let base = type_name
        .rsplit('.')
        .next()
        .unwrap_or(type_name)
        .split('<')
        .next()
        .unwrap_or(type_name)
        .trim();
    matches!(
        base,
        "LocalDate" | "LocalTime" | "LocalDateTime" | "OffsetDateTime" | "ZonedDateTime"
    )
}

fn java_type_is_duration(type_name: Option<&str>) -> bool {
    let Some(type_name) = type_name else {
        return false;
    };
    let base = type_name
        .rsplit('.')
        .next()
        .unwrap_or(type_name)
        .split('<')
        .next()
        .unwrap_or(type_name)
        .trim();
    base == "Duration"
}

fn java_type_is_zone_id(type_name: Option<&str>) -> bool {
    let Some(type_name) = type_name else {
        return false;
    };
    let base = type_name
        .rsplit('.')
        .next()
        .unwrap_or(type_name)
        .split('<')
        .next()
        .unwrap_or(type_name)
        .trim();
    matches!(base, "ZoneId" | "ZoneOffset")
}

fn java_instant_method_name(method: &str) -> Option<&'static str> {
    Some(match method {
        "compareTo" => "__java_instant_compare_to",
        "equals" => "__java_instant_equals",
        "toString" => "__java_instant_to_string",
        "hashCode" => "__java_instant_hash_code",
        _ => return None,
    })
}

fn java_time_method_name(method: &str) -> Option<&'static str> {
    Some(match method {
        "compareTo" => "__java_instant_compare_to",
        "equals" | "isEqual" => "__java_instant_equals",
        "isBefore" => "__java_instant_is_before",
        "isAfter" => "__java_instant_is_after",
        "hashCode" => "__java_instant_hash_code",
        "toString" => "__java_time_to_string",
        _ => return None,
    })
}

fn java_duration_method_name(method: &str) -> Option<&'static str> {
    Some(match method {
        "plusHours" => "__java_duration_plus_hours",
        "minusMinutes" => "__java_duration_minus_minutes",
        _ => return None,
    })
}

fn java_zone_method_name(method: &str) -> Option<&'static str> {
    Some(match method {
        "compareTo" => "__java_zone_compare_to",
        "hashCode" => "__java_zone_hash_code",
        _ => return None,
    })
}

fn java_uuid_method_name(method: &str) -> Option<&'static str> {
    Some(match method {
        "compareTo" => "__java_uuid_compare_to",
        "hashCode" => "__java_uuid_hash_code",
        _ => return None,
    })
}

fn java_bitset_method_name(method: &str) -> Option<&'static str> {
    Some(match method {
        "set" => "__java_bitset_set",
        "get" => "__java_bitset_get",
        "clear" => "__java_bitset_clear",
        "flip" => "__java_bitset_flip",
        "cardinality" => "__java_bitset_cardinality",
        "length" => "__java_bitset_length",
        "size" => "__java_bitset_size",
        "isEmpty" => "__java_bitset_is_empty",
        "nextSetBit" => "__java_bitset_next_set_bit",
        "nextClearBit" => "__java_bitset_next_clear_bit",
        "previousSetBit" => "__java_bitset_previous_set_bit",
        "previousClearBit" => "__java_bitset_previous_clear_bit",
        "and" => "__java_bitset_and",
        "or" => "__java_bitset_or",
        "xor" => "__java_bitset_xor",
        "andNot" => "__java_bitset_and_not",
        "intersects" => "__java_bitset_intersects",
        "equals" => "__java_bitset_equals",
        "clone" => "__java_bitset_clone",
        "stream" => "__java_bitset_stream",
        "toLongArray" | "toByteArray" => "__java_bitset_to_array",
        "toString" => "__java_bitset_to_string",
        "hashCode" => "__java_bitset_hash_code",
        _ => return None,
    })
}

fn java_bigint_method_name(method: &str) -> Option<&'static str> {
    match method {
        "toString" => Some("__java_bigint_to_string"),
        "add" => Some("__java_bigint_add"),
        "subtract" => Some("__java_bigint_subtract"),
        "multiply" => Some("__java_bigint_multiply"),
        "mod" => Some("__java_bigint_mod"),
        "gcd" => Some("__java_bigint_gcd"),
        "pow" => Some("__java_bigint_pow"),
        "compareTo" => Some("__java_bigint_compare_to"),
        "negate" => Some("__java_bigint_negate"),
        "abs" => Some("__java_bigint_abs"),
        "signum" => Some("__java_bigint_signum"),
        "max" => Some("__java_bigint_max"),
        "min" => Some("__java_bigint_min"),
        "bitLength" => Some("__java_bigint_bit_length"),
        "testBit" => Some("__java_bigint_test_bit"),
        "shiftLeft" => Some("__java_bigint_shift_left"),
        "shiftRight" => Some("__java_bigint_shift_right"),
        "and" => Some("__java_bigint_and"),
        "or" => Some("__java_bigint_or"),
        "xor" => Some("__java_bigint_xor"),
        "not" => Some("__java_bigint_not"),
        "isProbablePrime" => Some("__java_bigint_is_probable_prime"),
        "nextProbablePrime" => Some("__java_bigint_next_probable_prime"),
        _ => None,
    }
}

fn java_bigdecimal_method_name(method: &str, argc: usize) -> Option<&'static str> {
    match method {
        "toString" => Some("__j_bd_to_string"),
        "toPlainString" => Some("__j_bd_to_plain_string"),
        "add" => Some("__j_bd_add"),
        "subtract" => Some("__j_bd_subtract"),
        "multiply" => Some("__j_bd_multiply"),
        "divide" if argc == 1 => Some("__j_bd_divide"),
        "divide" if argc >= 3 => Some("__j_bd_divide_scale"),
        "scale" => Some("__j_bd_scale"),
        "compareTo" => Some("__j_bd_compare_to"),
        "stripTrailingZeros" => Some("__j_bd_strip"),
        "negate" => Some("__j_bd_negate"),
        "abs" => Some("__j_bd_abs"),
        "plus" => Some("__j_bd_plus"),
        "setScale" => Some("__j_bd_set_scale"),
        "movePointRight" => Some("__j_bd_move_right"),
        "movePointLeft" => Some("__j_bd_move_left"),
        "signum" => Some("__j_bd_signum"),
        "unscaledValue" => Some("__j_bd_unscaled"),
        "precision" => Some("__j_bd_precision"),
        "max" => Some("__j_bd_max"),
        "min" => Some("__j_bd_min"),
        "remainder" => Some("__j_bd_remainder"),
        "equals" => Some("__j_bd_equals"),
        _ => None,
    }
}

fn java_decimal_format_method_name(method: &str) -> Option<&'static str> {
    match method {
        "format" => Some("__j_df_format"),
        "parse" => Some("__j_df_parse"),
        "setMinimumFractionDigits" => Some("__j_df_set_min_frac"),
        "setMaximumFractionDigits" => Some("__j_df_set_max_frac"),
        "getMinimumFractionDigits" => Some("__j_df_min_frac"),
        "getMaximumFractionDigits" => Some("__j_df_max_frac"),
        "setGroupingUsed" => Some("__j_df_set_grouping"),
        "isGroupingUsed" => Some("__j_df_grouping"),
        "applyPattern" => Some("__j_df_apply_pattern"),
        "toPattern" => Some("__j_df_pattern"),
        "setDecimalSeparatorAlwaysShown" => Some("__j_df_set_decimal_always"),
        "getMultiplier" => Some("__j_df_multiplier"),
        "setMultiplier" => Some("__j_df_set_multiplier"),
        "setParseIntegerOnly" => Some("__j_df_set_parse_integer"),
        "clone" => Some("__j_df_clone"),
        "equals" => Some("__j_df_equals"),
        _ => None,
    }
}

fn java_stream_builder_receiver(expr: &Expression) -> bool {
    let ExprKind::Call { callee, .. } = &expr.kind else {
        return false;
    };
    matches!(
        &callee.kind,
        ExprKind::Ident(name)
            if matches!(
                name.as_str(),
                "IntStream.builder"
                    | "LongStream.builder"
                    | "DoubleStream.builder"
                    | "Stream.builder"
                    | "__j_stream_builder_add"
            )
    )
}

fn java_args_are_copy_sign_negative_zero(args: &[Argument]) -> bool {
    args.len() == 2
        && java_expr_is_zero_f64(&args[0].value)
        && java_expr_is_negative_f64(&args[1].value)
}

fn java_expr_is_zero_f64(expr: &Expression) -> bool {
    matches!(expr.kind, ExprKind::Lit(Literal::Float(value)) if value == 0.0)
}

fn java_expr_is_negative_f64(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Lit(Literal::Float(value)) => *value < 0.0,
        ExprKind::Lit(Literal::Int(value)) => *value < 0,
        ExprKind::Unary { expr, .. } => {
            java_expr_is_zero_f64(expr)
                || matches!(expr.kind, ExprKind::Lit(Literal::Float(value)) if value > 0.0)
        }
        _ => false,
    }
}

fn java_expr_is_bigint(
    expr: &Expression,
    locals: &std::collections::HashMap<String, String>,
) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => locals
            .get(name)
            .is_some_and(|type_hint| java_type_simple_name(type_hint) == "BigInteger"),
        ExprKind::New { class, .. } => match &class.kind {
            ExprKind::Ident(name) => java_type_simple_name(name) == "BigInteger",
            ExprKind::Member { .. } => {
                java_qualified_static_type(class).is_some_and(|name| name == "BigInteger")
            }
            _ => false,
        },
        ExprKind::Call { callee, .. } => match &callee.kind {
            ExprKind::Ident(name) => {
                name.starts_with("__java_bigint") || name == "BigInteger.valueOf"
            }
            _ => false,
        },
        _ => java_bigint_constant_replacement(expr).is_some(),
    }
}

fn java_expr_is_bigdecimal(
    expr: &Expression,
    locals: &std::collections::HashMap<String, String>,
) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => locals
            .get(name)
            .is_some_and(|type_hint| java_type_simple_name(type_hint) == "BigDecimal"),
        ExprKind::New { class, .. } => match &class.kind {
            ExprKind::Ident(name) => java_type_simple_name(name) == "BigDecimal",
            ExprKind::Member { .. } => java_qualified_static_type(class)
                .is_some_and(|name| java_type_simple_name(&name) == "BigDecimal"),
            _ => false,
        },
        ExprKind::Call { callee, .. } => {
            matches!(&callee.kind, ExprKind::Ident(name) if java_bigdecimal_function_returns_bigdecimal(name))
        }
        _ => java_bigdecimal_constant_replacement(expr).is_some(),
    }
}

fn java_bigdecimal_function_returns_bigdecimal(name: &str) -> bool {
    matches!(
        name,
        "__j_bd_new"
            | "__j_bd_box"
            | "__j_bd_add"
            | "__j_bd_subtract"
            | "__j_bd_multiply"
            | "__j_bd_divide"
            | "__j_bd_divide_scale"
            | "__j_bd_set_scale"
            | "__j_bd_strip"
            | "__j_bd_negate"
            | "__j_bd_abs"
            | "__j_bd_plus"
            | "__j_bd_move_right"
            | "__j_bd_move_left"
            | "__j_bd_max"
            | "__j_bd_min"
            | "__j_bd_remainder"
    )
}

fn java_bigint_constant_replacement(expr: &Expression) -> Option<Expression> {
    if let ExprKind::Member { object, field, .. } = &expr.kind {
        let is_bigint_type = java_qualified_static_type(object)
            .is_some_and(|name| name == "BigInteger")
            || java_expr_dotted_name(object).as_deref() == Some("java.math.BigInteger");
        if is_bigint_type {
            let value = match field.as_str() {
                "ZERO" => "0",
                "ONE" => "1",
                "TEN" => "10",
                _ => return None,
            };
            return Some(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__java_bigint")),
                args: vec![Argument::positional(Expression::int(
                    value.parse::<i64>().unwrap_or(0),
                ))],
                optional: false,
            }));
        }
    }
    None
}

fn java_bigdecimal_constant_replacement(expr: &Expression) -> Option<Expression> {
    if let ExprKind::Member { object, field, .. } = &expr.kind {
        let is_bigdecimal_type = java_qualified_static_type(object)
            .is_some_and(|name| java_type_simple_name(&name) == "BigDecimal")
            || java_expr_dotted_name(object)
                .as_deref()
                .is_some_and(|name| name == "java.math.BigDecimal");
        if is_bigdecimal_type {
            let value = match field.as_str() {
                "ZERO" => "0",
                "ONE" => "1",
                "TEN" => "10",
                _ => return None,
            };
            return Some(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__j_bd_new")),
                args: vec![Argument::positional(Expression::string(value))],
                optional: false,
            }));
        }
    }
    None
}

fn java_type_simple_name(type_name: &str) -> &str {
    let erased = common_generics::generic_base_name(type_name.trim());
    erased.rsplit('.').next().unwrap_or(erased)
}

fn java_type_base_simple_name(type_name: &str) -> &str {
    java_type_simple_name(type_name).trim()
}

fn java_print_arg_needs_tostring(
    arg: &Expression,
    tostring_classes: &std::collections::HashSet<String>,
    enum_values: &std::collections::HashMap<String, Vec<String>>,
    current_class: Option<&str>,
    locals: &std::collections::HashMap<String, String>,
) -> bool {
    if java_expr_has_user_tostring(arg, tostring_classes, current_class, locals) {
        return true;
    }
    java_expr_enum_type(arg, enum_values, current_class, locals).is_some()
}

fn java_tostring_call(receiver: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(receiver),
            field: "tostring".to_string(),
            null_safe: false,
        })),
        args: vec![],
        optional: false,
    })
}

fn java_expr_enum_type(
    expr: &Expression,
    enum_values: &std::collections::HashMap<String, Vec<String>>,
    current_class: Option<&str>,
    locals: &std::collections::HashMap<String, String>,
) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => {
            if let Some(type_hint) = locals.get(name) {
                let simple = java_type_simple_name(type_hint);
                if enum_values.contains_key(simple) {
                    return Some(simple.to_string());
                }
            }
            None
        }
        ExprKind::Member { object, field, .. } => {
            if let ExprKind::Ident(type_name) = &object.kind {
                let simple = java_type_simple_name(type_name);
                if enum_values
                    .get(simple)
                    .is_some_and(|members| members.iter().any(|m| m == field))
                {
                    return Some(simple.to_string());
                }
            }
            None
        }
        ExprKind::Index { object, .. } => {
            if let ExprKind::Call { callee, args, .. } = &object.kind {
                if args.is_empty() {
                    if let ExprKind::Member {
                        object: enum_object,
                        field,
                        ..
                    } = &callee.kind
                    {
                        if field == "values" {
                            if let ExprKind::Ident(type_name) = &enum_object.kind {
                                let simple = java_type_simple_name(type_name);
                                if enum_values.contains_key(simple) {
                                    return Some(simple.to_string());
                                }
                            }
                        }
                    }
                }
            }
            None
        }
        ExprKind::Call { callee, args, .. } => match &callee.kind {
            ExprKind::Member { object, field, .. } => {
                if field == "__j_enum_value_of" {
                    if let ExprKind::Ident(type_name) = &object.kind {
                        let simple = java_type_simple_name(type_name);
                        if enum_values.contains_key(simple) {
                            return Some(simple.to_string());
                        }
                    }
                }
                let receiver_type = match &object.kind {
                    ExprKind::Ident(type_name) if enum_values.contains_key(type_name.as_str()) => {
                        Some(type_name.clone())
                    }
                    _ => java_expr_enum_type(object, enum_values, current_class, locals),
                };
                if let Some(receiver_type) = receiver_type {
                    if let Some(return_type) = java_class_method_return_type(&receiver_type, field)
                    {
                        let simple = java_type_simple_name(&return_type);
                        if enum_values.contains_key(simple) {
                            return Some(simple.to_string());
                        }
                    }
                }
                if args.is_empty() {
                    let dotted = java_expr_dotted_name(object)?;
                    let key = format!("{dotted}.{field}()");
                    if let Some(type_hint) = locals.get(&key) {
                        let simple = java_type_simple_name(type_hint);
                        if enum_values.contains_key(simple) {
                            return Some(simple.to_string());
                        }
                    }
                }
                None
            }
            ExprKind::Ident(name) if args.is_empty() => {
                let key = format!("{name}()");
                if let Some(type_hint) = locals.get(&key) {
                    let simple = java_type_simple_name(type_hint);
                    if enum_values.contains_key(simple) {
                        return Some(simple.to_string());
                    }
                }
                if let Some(class_name) = current_class {
                    let key = format!("{class_name}.{name}()");
                    if let Some(type_hint) = locals.get(&key) {
                        let simple = java_type_simple_name(type_hint);
                        if enum_values.contains_key(simple) {
                            return Some(simple.to_string());
                        }
                    }
                }
                None
            }
            _ => None,
        },
        _ => None,
    }
}

fn java_class_method_return_type(class_name: &str, method_name: &str) -> Option<String> {
    let simple = java_type_simple_name(class_name);
    JAVA_REFLECTION_CLASSES.with(|classes| {
        classes.borrow().get(simple).and_then(|meta| {
            meta.methods
                .iter()
                .find(|method| method.name == method_name)
                .and_then(|method| method.return_type.clone())
        })
    })
}

fn java_class_field_type(class_name: &str, field_name: &str) -> Option<String> {
    let simple = java_type_simple_name(class_name);
    JAVA_REFLECTION_CLASSES.with(|classes| {
        classes.borrow().get(simple).and_then(|meta| {
            meta.fields
                .iter()
                .find(|field| field.name == field_name)
                .and_then(|field| field.type_name.clone())
        })
    })
}

fn java_expr_has_user_tostring(
    expr: &Expression,
    tostring_classes: &std::collections::HashSet<String>,
    current_class: Option<&str>,
    locals: &std::collections::HashMap<String, String>,
) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => locals
            .get(name)
            .is_some_and(|type_hint| tostring_classes.contains(type_hint)),
        ExprKind::This => current_class.is_some_and(|name| tostring_classes.contains(name)),
        ExprKind::New { class, .. } => match &class.kind {
            ExprKind::Ident(name) => tostring_classes.contains(name),
            _ => false,
        },
        _ => false,
    }
}

fn normalize_java_class_tree(body: &mut [Statement]) {
    use std::collections::HashMap;

    let mut class_members = HashMap::new();
    let mut class_parents = HashMap::new();
    collect_java_class_member_names(body, &mut class_members, &mut class_parents);
    install_java_interface_default_methods(body, &class_members, &class_parents);
    class_members.clear();
    class_parents.clear();
    collect_java_class_member_names(body, &mut class_members, &mut class_parents);
    normalize_java_class_tree_with_members(body, &class_members, &class_parents);
    let mut locals = std::collections::HashSet::new();
    lower_java_anonymous_class_captures_tree(body, &mut locals);
    normalize_java_anonymous_class_tree(body, &class_members);
}

fn inject_java_static_initializer_calls(body: &mut Vec<Statement>) {
    let mut class_names = Vec::new();
    collect_java_static_initializer_classes(body, &mut class_names);
    if class_names.is_empty() {
        return;
    }
    let calls = class_names.into_iter().map(|class_name| {
        Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident(&class_name)),
                field: "__static_init_block__".to_string(),
                null_safe: false,
            })),
            args: vec![],
            optional: false,
        })))
    });
    body.extend(calls);
}

fn collect_java_static_initializer_classes(body: &[Statement], out: &mut Vec<String>) {
    for stmt in body {
        match &stmt.kind {
            StmtKind::ClassDecl { name, members, .. } => {
                if members.iter().any(|member| {
                    matches!(
                        member,
                        ClassMember::Method(method)
                            if matches!(&method.kind, StmtKind::FunctionDecl { name, .. } if name == "__static_init_block__")
                    )
                }) {
                    out.push(name.clone());
                }
                for member in members {
                    if let ClassMember::NestedType(nested) = member {
                        collect_java_static_initializer_classes(std::slice::from_ref(nested), out);
                    }
                }
            }
            StmtKind::EnumDecl { body_members, .. } => {
                for member in body_members {
                    if let ClassMember::NestedType(nested) = member {
                        collect_java_static_initializer_classes(std::slice::from_ref(nested), out);
                    }
                }
            }
            StmtKind::Block(stmts) => collect_java_static_initializer_classes(stmts, out),
            _ => {}
        }
    }
}

fn collect_java_class_member_names(
    body: &[Statement],
    out: &mut std::collections::HashMap<String, JavaClassMemberNames>,
    parents_out: &mut std::collections::HashMap<String, Vec<String>>,
) {
    for stmt in body {
        match &stmt.kind {
            StmtKind::ClassDecl {
                name,
                parents,
                interfaces,
                members,
                modifiers,
                ..
            } => {
                let is_interface =
                    JAVA_INTERFACE_NAMES.with(|names| names.borrow().contains(name.as_str()));
                let is_enum =
                    JAVA_ENUM_VALUES.with(|values| values.borrow().contains_key(name.as_str()));
                JAVA_REFLECTION_CLASSES.with(|classes| {
                    classes.borrow_mut().insert(
                        name.clone(),
                        java_reflection_meta(
                            name,
                            parents.first().cloned(),
                            interfaces.clone(),
                            members,
                            is_interface,
                            is_enum,
                            java_class_modifier_bits(modifiers),
                        ),
                    );
                });
                out.insert(name.clone(), JavaClassMemberNames::from_members(members));
                let mut inherited = parents.clone();
                inherited.extend(interfaces.iter().cloned());
                parents_out.insert(name.clone(), inherited);
                for member in members {
                    if let ClassMember::NestedType(nested) = member {
                        collect_java_class_member_names(
                            std::slice::from_ref(nested),
                            out,
                            parents_out,
                        );
                    }
                }
            }
            // Enums are class-shaped (JLS §8.9): their body members get the
            // same bare field/method qualification as class members.
            StmtKind::EnumDecl {
                name, body_members, ..
            } => {
                out.insert(
                    name.clone(),
                    JavaClassMemberNames::from_members(body_members),
                );
                for member in body_members {
                    if let ClassMember::NestedType(nested) = member {
                        collect_java_class_member_names(
                            std::slice::from_ref(nested),
                            out,
                            parents_out,
                        );
                    }
                }
            }
            StmtKind::Block(stmts) => collect_java_class_member_names(stmts, out, parents_out),
            _ => {}
        }
    }
}

fn normalize_java_class_tree_with_members(
    body: &mut [Statement],
    class_members: &std::collections::HashMap<String, JavaClassMemberNames>,
    class_parents: &std::collections::HashMap<String, Vec<String>>,
) {
    for stmt in body {
        match &mut stmt.kind {
            StmtKind::ClassDecl {
                name,
                parents,
                interfaces,
                members,
                ..
            } => {
                let mut names = class_members.get(name).cloned().unwrap_or_default();
                let mut inherited = parents.clone();
                inherited.extend(interfaces.iter().cloned());
                merge_java_inherited_member_names(
                    &mut names,
                    &inherited,
                    class_members,
                    class_parents,
                );
                normalize_java_class_members(members, name, &names, class_members);
                for member in members {
                    if let ClassMember::NestedType(nested) = member {
                        normalize_java_class_tree_with_members(
                            std::slice::from_mut(nested),
                            class_members,
                            class_parents,
                        );
                    }
                }
            }
            StmtKind::EnumDecl {
                name, body_members, ..
            } => {
                let names = class_members.get(name).cloned().unwrap_or_default();
                normalize_java_class_members(body_members, name, &names, class_members);
                for member in body_members {
                    if let ClassMember::NestedType(nested) = member {
                        normalize_java_class_tree_with_members(
                            std::slice::from_mut(nested),
                            class_members,
                            class_parents,
                        );
                    }
                }
            }
            StmtKind::Block(stmts) => {
                normalize_java_class_tree_with_members(stmts, class_members, class_parents)
            }
            _ => {}
        }
    }
}

fn merge_java_inherited_member_names(
    names: &mut JavaClassMemberNames,
    parents: &[String],
    class_members: &std::collections::HashMap<String, JavaClassMemberNames>,
    class_parents: &std::collections::HashMap<String, Vec<String>>,
) {
    let mut stack: Vec<String> = parents.to_vec();
    let mut seen = std::collections::HashSet::new();
    while let Some(parent) = stack.pop() {
        if !seen.insert(parent.clone()) {
            continue;
        }
        if let Some(parent_names) = class_members.get(&parent) {
            names.fields.extend(parent_names.fields.iter().cloned());
            names.field_types.extend(
                parent_names
                    .field_types
                    .iter()
                    .map(|(name, ty)| (name.clone(), ty.clone())),
            );
            names.methods.extend(parent_names.methods.iter().cloned());
            names
                .static_methods
                .extend(parent_names.static_methods.iter().cloned());
        }
        if let Some(parent_parents) = class_parents.get(&parent) {
            stack.extend(parent_parents.iter().cloned());
        }
    }
}

fn install_java_interface_default_methods(
    body: &mut [Statement],
    class_members: &std::collections::HashMap<String, JavaClassMemberNames>,
    class_parents: &std::collections::HashMap<String, Vec<String>>,
) {
    for stmt in body {
        match &mut stmt.kind {
            StmtKind::ClassDecl {
                name,
                parents,
                interfaces,
                members,
                ..
            } => {
                let is_interface = JAVA_INTERFACE_NAMES.with(|names| names.borrow().contains(name));
                if !is_interface {
                    let mut implemented = parents.clone();
                    implemented.extend(interfaces.iter().cloned());
                    let mut defaults = Vec::new();
                    collect_java_interface_default_methods(
                        &implemented,
                        class_members,
                        class_parents,
                        &mut std::collections::HashSet::new(),
                        &mut defaults,
                    );
                    append_missing_java_interface_defaults(members, defaults);
                }
                for member in members {
                    if let ClassMember::NestedType(nested) = member {
                        install_java_interface_default_methods(
                            std::slice::from_mut(nested),
                            class_members,
                            class_parents,
                        );
                    }
                }
            }
            StmtKind::Block(stmts) => {
                install_java_interface_default_methods(stmts, class_members, class_parents);
            }
            _ => {}
        }
    }
}

fn collect_java_interface_default_methods(
    interfaces: &[String],
    class_members: &std::collections::HashMap<String, JavaClassMemberNames>,
    class_parents: &std::collections::HashMap<String, Vec<String>>,
    seen: &mut std::collections::HashSet<String>,
    out: &mut Vec<ClassMember>,
) {
    for interface in interfaces {
        let is_interface = JAVA_INTERFACE_NAMES.with(|names| names.borrow().contains(interface));
        if !is_interface || !seen.insert(interface.clone()) {
            continue;
        }
        if let Some(parents) = class_parents.get(interface) {
            collect_java_interface_default_methods(
                parents,
                class_members,
                class_parents,
                seen,
                out,
            );
        }
        if let Some(names) = class_members.get(interface) {
            for method in &names.default_methods {
                let original = (**method).clone();
                if let Some(method_name) = java_function_decl_name(&original) {
                    let mut hidden = original.clone();
                    java_rename_function_decl(
                        &mut hidden,
                        &java_interface_default_method_name(interface, &method_name),
                    );
                    out.push(ClassMember::Method(Box::new(hidden)));
                }
                out.push(ClassMember::Method(Box::new(original)));
            }
        }
    }
}

fn append_missing_java_interface_defaults(
    members: &mut Vec<ClassMember>,
    defaults: Vec<ClassMember>,
) {
    if defaults.is_empty() {
        return;
    }
    let mut existing = java_instance_method_names(members);
    for default_method in defaults {
        let Some(name) = java_class_method_name(&default_method) else {
            continue;
        };
        if existing.insert(name) {
            members.push(default_method);
        }
    }
}

fn java_class_method_name(member: &ClassMember) -> Option<String> {
    let ClassMember::Method(func) = member else {
        return None;
    };
    java_function_decl_name(func).and_then(|name| {
        let StmtKind::FunctionDecl { modifiers, .. } = &func.kind else {
            return None;
        };
        (!modifiers.is_static).then_some(name)
    })
}

fn java_function_decl_name(func: &Statement) -> Option<String> {
    let StmtKind::FunctionDecl {
        name, modifiers, ..
    } = &func.kind
    else {
        return None;
    };
    (!modifiers.is_static).then(|| name.clone())
}

fn java_rename_function_decl(func: &mut Statement, renamed: &str) {
    if let StmtKind::FunctionDecl { name, .. } = &mut func.kind {
        *name = renamed.to_string();
    }
}

fn java_interface_default_method_name(interface: &str, method: &str) -> String {
    let interface = java_type_simple_name(interface)
        .chars()
        .map(|ch| if ch.is_ascii_alphanumeric() { ch } else { '_' })
        .collect::<String>();
    format!("__java_default_{interface}_{method}")
}

fn normalize_java_anonymous_class_tree(
    body: &mut [Statement],
    class_members: &std::collections::HashMap<String, JavaClassMemberNames>,
) {
    for stmt in body {
        normalize_java_anonymous_class_stmt(stmt, class_members);
    }
}

fn lower_java_anonymous_class_captures_tree(
    body: &mut [Statement],
    locals: &mut std::collections::HashSet<String>,
) {
    for stmt in body {
        lower_java_anonymous_class_captures_stmt(stmt, locals);
    }
}

fn lower_java_anonymous_class_captures_members(
    members: &mut [ClassMember],
    locals: &std::collections::HashSet<String>,
) {
    for member in members {
        match member {
            ClassMember::Field { init, .. } => {
                if let Some(init) = init {
                    lower_java_anonymous_class_captures_expr(init, &mut locals.clone());
                }
            }
            ClassMember::Constructor { params, body, .. } => {
                let mut member_locals = locals.clone();
                for param in params {
                    member_locals.insert(param.name.clone());
                }
                lower_java_anonymous_class_captures_tree(body, &mut member_locals);
            }
            ClassMember::Method(method) => {
                if let StmtKind::FunctionDecl { params, body, .. } = &mut method.kind {
                    let mut member_locals = locals.clone();
                    for param in params {
                        member_locals.insert(param.name.clone());
                    }
                    lower_java_anonymous_class_captures_tree(body, &mut member_locals);
                }
            }
            ClassMember::NestedType(nested) => {
                lower_java_anonymous_class_captures_stmt(nested, &mut locals.clone());
            }
            _ => {}
        }
    }
}

fn lower_java_anonymous_class_captures_stmt(
    stmt: &mut Statement,
    locals: &mut std::collections::HashSet<String>,
) {
    match &mut stmt.kind {
        StmtKind::Expr(expr)
        | StmtKind::Return(Some(expr))
        | StmtKind::Throw {
            expr: Some(expr), ..
        } => {
            lower_java_anonymous_class_captures_expr(expr, locals);
        }
        StmtKind::Block(stmts) | StmtKind::NamespaceDecl { body: stmts, .. } => {
            lower_java_anonymous_class_captures_tree(stmts, &mut locals.clone());
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    lower_java_anonymous_class_captures_expr(init, locals);
                }
                collect_binding_names(&decl.pattern, locals);
            }
        }
        StmtKind::FunctionDecl { params, body, .. } => {
            let mut fn_locals = locals.clone();
            for param in params {
                fn_locals.insert(param.name.clone());
            }
            lower_java_anonymous_class_captures_tree(body, &mut fn_locals);
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::EnumDecl {
            body_members: members,
            ..
        } => lower_java_anonymous_class_captures_members(members, locals),
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            lower_java_anonymous_class_captures_expr(cond, locals);
            lower_java_anonymous_class_captures_tree(then_body, &mut locals.clone());
            for (elif_cond, elif_body) in elifs {
                lower_java_anonymous_class_captures_expr(elif_cond, locals);
                lower_java_anonymous_class_captures_tree(elif_body, &mut locals.clone());
            }
            if let Some(else_body) = else_body {
                lower_java_anonymous_class_captures_tree(else_body, &mut locals.clone());
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            let mut loop_locals = locals.clone();
            if let Some(init) = init {
                lower_java_anonymous_class_captures_stmt(init, &mut loop_locals);
            }
            if let Some(cond) = cond {
                lower_java_anonymous_class_captures_expr(cond, &mut loop_locals);
            }
            if let Some(update) = update {
                lower_java_anonymous_class_captures_expr(update, &mut loop_locals);
            }
            lower_java_anonymous_class_captures_tree(body, &mut loop_locals);
        }
        StmtKind::ForIn {
            var,
            key,
            iter,
            body,
            else_body,
            ..
        } => {
            lower_java_anonymous_class_captures_expr(iter, locals);
            let mut loop_locals = locals.clone();
            loop_locals.insert(var.clone());
            if let Some(key) = key {
                loop_locals.insert(key.clone());
            }
            lower_java_anonymous_class_captures_tree(body, &mut loop_locals);
            if let Some(else_body) = else_body {
                lower_java_anonymous_class_captures_tree(else_body, &mut locals.clone());
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            lower_java_anonymous_class_captures_expr(cond, locals);
            lower_java_anonymous_class_captures_tree(body, &mut locals.clone());
            if let Some(else_body) = else_body {
                lower_java_anonymous_class_captures_tree(else_body, &mut locals.clone());
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            lower_java_anonymous_class_captures_tree(body, &mut locals.clone());
            lower_java_anonymous_class_captures_expr(cond, locals);
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                lower_java_anonymous_class_captures_expr(target, locals);
            }
            lower_java_anonymous_class_captures_expr(value, locals);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            lower_java_anonymous_class_captures_expr(target, locals);
            lower_java_anonymous_class_captures_expr(value, locals);
        }
        _ => {}
    }
}

fn lower_java_anonymous_class_captures_expr(
    expr: &mut Expression,
    locals: &mut std::collections::HashSet<String>,
) {
    match &mut expr.kind {
        ExprKind::New { class, args } => {
            for arg in args.iter_mut() {
                lower_java_anonymous_class_captures_expr(&mut arg.value, locals);
            }
            lower_java_anonymous_class_captures_expr(class, locals);
            if let ExprKind::ClassExpr {
                parent, members, ..
            } = &mut class.kind
            {
                inject_java_anonymous_captures(parent.as_deref(), members, args, locals);
            }
        }
        ExprKind::Binary { left, right, .. } => {
            lower_java_anonymous_class_captures_expr(left, locals);
            lower_java_anonymous_class_captures_expr(right, locals);
        }
        ExprKind::Unary { expr: inner, .. }
        | ExprKind::Spread(inner)
        | ExprKind::Await(inner)
        | ExprKind::YieldFrom(inner)
        | ExprKind::Void(inner)
        | ExprKind::Delete(inner)
        | ExprKind::TypeOf(inner)
        | ExprKind::RefLoad(inner) => lower_java_anonymous_class_captures_expr(inner, locals),
        ExprKind::Yield(Some(inner)) => lower_java_anonymous_class_captures_expr(inner, locals),
        ExprKind::Ternary { cond, then, else_ } => {
            lower_java_anonymous_class_captures_expr(cond, locals);
            lower_java_anonymous_class_captures_expr(then, locals);
            lower_java_anonymous_class_captures_expr(else_, locals);
        }
        ExprKind::Member { object, .. } => lower_java_anonymous_class_captures_expr(object, locals),
        ExprKind::Index { object, index, .. } => {
            lower_java_anonymous_class_captures_expr(object, locals);
            lower_java_anonymous_class_captures_expr(index, locals);
        }
        ExprKind::Call { callee, args, .. } => {
            lower_java_anonymous_class_captures_expr(callee, locals);
            for arg in args {
                lower_java_anonymous_class_captures_expr(&mut arg.value, locals);
            }
        }
        ExprKind::Assign { target, value } => {
            lower_java_anonymous_class_captures_expr(target, locals);
            lower_java_anonymous_class_captures_expr(value, locals);
        }
        ExprKind::Array(elems) => {
            for elem in elems {
                lower_java_anonymous_class_captures_expr(&mut elem.value, locals);
            }
        }
        ExprKind::Tuple(elems) | ExprKind::Set(elems) | ExprKind::Sequence(elems) => {
            for elem in elems {
                lower_java_anonymous_class_captures_expr(elem, locals);
            }
        }
        ExprKind::Lambda { body, params, .. } => {
            let mut used = std::collections::HashSet::new();
            match body {
                LambdaBody::Expr(inner) => collect_java_lambda_used_idents_expr(inner, &mut used),
                LambdaBody::Block(stmts) => collect_java_lambda_used_idents_stmts(stmts, &mut used),
            }
            let mut inner_locals: std::collections::HashSet<String> =
                params.iter().map(|param| param.name.clone()).collect();
            if let LambdaBody::Block(stmts) = body {
                collect_java_lambda_declared_names(stmts, &mut inner_locals);
            }
            let mut lambda_locals = locals.clone();
            for param in params {
                lambda_locals.insert(param.name.clone());
            }
            let _captured_locals: Vec<String> = used
                .into_iter()
                .filter(|name| locals.contains(name) && !inner_locals.contains(name))
                .collect();
            match body {
                LambdaBody::Expr(inner) => {
                    lower_java_anonymous_class_captures_expr(inner, &mut lambda_locals)
                }
                LambdaBody::Block(stmts) => {
                    lower_java_anonymous_class_captures_tree(stmts, &mut lambda_locals)
                }
            }
        }
        ExprKind::ClassExpr { members, .. } => {
            lower_java_anonymous_class_captures_members(members, locals);
        }
        ExprKind::FunctionExpr(func) => lower_java_anonymous_class_captures_stmt(func, locals),
        _ => {}
    }
}

fn collect_java_lambda_declared_names(
    stmts: &[Statement],
    out: &mut std::collections::HashSet<String>,
) {
    for stmt in stmts {
        match &stmt.kind {
            StmtKind::VarDecl { declarations, .. } => {
                for decl in declarations {
                    collect_binding_names(&decl.pattern, out);
                }
            }
            StmtKind::Block(body) => collect_java_lambda_declared_names(body, out),
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                collect_java_lambda_declared_names(then_body, out);
                for (_, elif_body) in elifs {
                    collect_java_lambda_declared_names(elif_body, out);
                }
                if let Some(else_body) = else_body {
                    collect_java_lambda_declared_names(else_body, out);
                }
            }
            StmtKind::For { init, body, .. } => {
                if let Some(init) = init {
                    collect_java_lambda_declared_names(std::slice::from_ref(init), out);
                }
                collect_java_lambda_declared_names(body, out);
            }
            StmtKind::ForIn { var, key, body, .. } => {
                out.insert(var.clone());
                if let Some(key) = key {
                    out.insert(key.clone());
                }
                collect_java_lambda_declared_names(body, out);
            }
            StmtKind::While { body, .. } | StmtKind::DoWhile { body, .. } => {
                collect_java_lambda_declared_names(body, out);
            }
            _ => {}
        }
    }
}

fn collect_java_lambda_used_idents_stmts(
    stmts: &[Statement],
    out: &mut std::collections::HashSet<String>,
) {
    for stmt in stmts {
        match &stmt.kind {
            StmtKind::VarDecl { declarations, .. } => {
                for decl in declarations {
                    if let Some(init) = &decl.init {
                        collect_java_lambda_used_idents_expr(init, out);
                    }
                }
            }
            StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
                collect_java_lambda_used_idents_expr(expr, out);
            }
            StmtKind::Assign { targets, value } => {
                for target in targets {
                    collect_java_lambda_used_idents_expr(target, out);
                }
                collect_java_lambda_used_idents_expr(value, out);
            }
            StmtKind::CompoundAssign { target, value, .. } => {
                collect_java_lambda_used_idents_expr(target, out);
                collect_java_lambda_used_idents_expr(value, out);
            }
            StmtKind::Block(body) => collect_java_lambda_used_idents_stmts(body, out),
            StmtKind::If {
                cond,
                then_body,
                elifs,
                else_body,
            } => {
                collect_java_lambda_used_idents_expr(cond, out);
                collect_java_lambda_used_idents_stmts(then_body, out);
                for (elif_cond, elif_body) in elifs {
                    collect_java_lambda_used_idents_expr(elif_cond, out);
                    collect_java_lambda_used_idents_stmts(elif_body, out);
                }
                if let Some(else_body) = else_body {
                    collect_java_lambda_used_idents_stmts(else_body, out);
                }
            }
            StmtKind::While { cond, body, .. } | StmtKind::DoWhile { cond, body, .. } => {
                collect_java_lambda_used_idents_expr(cond, out);
                collect_java_lambda_used_idents_stmts(body, out);
            }
            StmtKind::For {
                init,
                cond,
                update,
                body,
            } => {
                if let Some(init) = init {
                    collect_java_lambda_used_idents_stmts(std::slice::from_ref(init), out);
                }
                if let Some(cond) = cond {
                    collect_java_lambda_used_idents_expr(cond, out);
                }
                if let Some(update) = update {
                    collect_java_lambda_used_idents_expr(update, out);
                }
                collect_java_lambda_used_idents_stmts(body, out);
            }
            StmtKind::ForIn { iter, body, .. } => {
                collect_java_lambda_used_idents_expr(iter, out);
                collect_java_lambda_used_idents_stmts(body, out);
            }
            StmtKind::Throw {
                expr: Some(expr), ..
            } => {
                collect_java_lambda_used_idents_expr(expr, out);
            }
            _ => {}
        }
    }
}

fn collect_java_lambda_used_idents_expr(
    expr: &Expression,
    out: &mut std::collections::HashSet<String>,
) {
    match &expr.kind {
        ExprKind::Ident(name) => {
            out.insert(name.clone());
        }
        ExprKind::Binary { left, right, .. } => {
            collect_java_lambda_used_idents_expr(left, out);
            collect_java_lambda_used_idents_expr(right, out);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Spread(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr)
        | ExprKind::RefLoad(expr) => collect_java_lambda_used_idents_expr(expr, out),
        ExprKind::Yield(Some(expr)) => collect_java_lambda_used_idents_expr(expr, out),
        ExprKind::Member { object, .. } => collect_java_lambda_used_idents_expr(object, out),
        ExprKind::Index { object, index, .. } => {
            collect_java_lambda_used_idents_expr(object, out);
            collect_java_lambda_used_idents_expr(index, out);
        }
        ExprKind::Call { callee, args, .. } => {
            collect_java_lambda_used_idents_expr(callee, out);
            for arg in args {
                collect_java_lambda_used_idents_expr(&arg.value, out);
            }
        }
        ExprKind::Assign { target, value } => {
            collect_java_lambda_used_idents_expr(target, out);
            collect_java_lambda_used_idents_expr(value, out);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            collect_java_lambda_used_idents_expr(cond, out);
            collect_java_lambda_used_idents_expr(then, out);
            collect_java_lambda_used_idents_expr(else_, out);
        }
        ExprKind::Array(elems) => {
            for elem in elems {
                collect_java_lambda_used_idents_expr(&elem.value, out);
            }
        }
        ExprKind::Tuple(elems) | ExprKind::Set(elems) | ExprKind::Sequence(elems) => {
            for elem in elems {
                collect_java_lambda_used_idents_expr(elem, out);
            }
        }
        ExprKind::New { class, args } => {
            collect_java_lambda_used_idents_expr(class, out);
            for arg in args {
                collect_java_lambda_used_idents_expr(&arg.value, out);
            }
        }
        ExprKind::Lambda { .. } => {}
        _ => {}
    }
}

fn inject_java_anonymous_captures(
    parent: Option<&Expression>,
    members: &mut Vec<ClassMember>,
    args: &mut Vec<Argument>,
    locals: &std::collections::HashSet<String>,
) {
    let member_names = JavaClassMemberNames::from_members(members);
    let mut used = std::collections::HashSet::new();
    collect_java_anonymous_member_idents(members, &member_names, &mut used);
    let mut captures: Vec<String> = used
        .into_iter()
        .filter(|name| locals.contains(name))
        .collect();
    captures.sort();
    captures.dedup();
    if captures.is_empty() {
        return;
    }

    let original_args: Vec<Expression> = args.iter().map(|arg| arg.value.clone()).collect();
    let original_arg_count = original_args.len();
    let capture_fields: Vec<(String, String)> = captures
        .iter()
        .map(|name| (name.clone(), format!("__java_capture_{name}")))
        .collect();

    for (_, field_name) in &capture_fields {
        members.insert(
            0,
            ClassMember::Field {
                name: field_name.clone(),
                type_hint: None,
                init: None,
                modifiers: Modifiers::default(),
                with_events: false,
                array_bounds: None,
            },
        );
    }

    rewrite_java_anonymous_capture_refs(members, &capture_fields);

    for capture in &captures {
        args.push(Argument::positional(Expression::ident(capture)));
    }

    let mut params = Vec::new();
    let mut base_args = Vec::new();
    for index in 0..original_arg_count {
        let name = format!("__java_super_arg_{index}");
        params.push(Param {
            name: name.clone(),
            type_hint: None,
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        });
        base_args.push(Expression::ident(&name));
    }
    for (_, field_name) in &capture_fields {
        params.push(Param {
            name: field_name.clone(),
            type_hint: None,
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        });
    }

    let body = capture_fields
        .iter()
        .map(|(_, field_name)| {
            Statement::new(StmtKind::Assign {
                targets: vec![Expression::new(ExprKind::Member {
                    object: Box::new(Expression::new(ExprKind::This)),
                    field: field_name.clone(),
                    null_safe: false,
                })],
                value: Expression::ident(field_name),
            })
        })
        .collect();

    members.insert(
        0,
        ClassMember::Constructor {
            name: None,
            params,
            body,
            base_args: (parent.is_some() && original_arg_count > 0).then_some(base_args),
            initializer_target: ConstructorInitializerTarget::Base,
            visibility: Visibility::Public,
        },
    );
}

fn collect_java_anonymous_member_idents(
    members: &[ClassMember],
    member_names: &JavaClassMemberNames,
    out: &mut std::collections::HashSet<String>,
) {
    for member in members {
        if let ClassMember::Method(method) = member {
            if let StmtKind::FunctionDecl { params, body, .. } = &method.kind {
                let mut locals: std::collections::HashSet<String> =
                    params.iter().map(|param| param.name.clone()).collect();
                collect_java_capture_idents_stmts(body, &mut locals, member_names, out);
            }
        }
    }
}

fn collect_java_capture_idents_stmts(
    stmts: &[Statement],
    locals: &mut std::collections::HashSet<String>,
    member_names: &JavaClassMemberNames,
    out: &mut std::collections::HashSet<String>,
) {
    for stmt in stmts {
        match &stmt.kind {
            StmtKind::VarDecl { declarations, .. } => {
                for decl in declarations {
                    if let Some(init) = &decl.init {
                        collect_java_capture_idents_expr(init, locals, member_names, out);
                    }
                    collect_binding_names(&decl.pattern, locals);
                }
            }
            StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
                collect_java_capture_idents_expr(expr, locals, member_names, out);
            }
            StmtKind::Block(body) => {
                collect_java_capture_idents_stmts(body, &mut locals.clone(), member_names, out);
            }
            StmtKind::If {
                cond,
                then_body,
                elifs,
                else_body,
            } => {
                collect_java_capture_idents_expr(cond, locals, member_names, out);
                collect_java_capture_idents_stmts(
                    then_body,
                    &mut locals.clone(),
                    member_names,
                    out,
                );
                for (elif_cond, elif_body) in elifs {
                    collect_java_capture_idents_expr(elif_cond, locals, member_names, out);
                    collect_java_capture_idents_stmts(
                        elif_body,
                        &mut locals.clone(),
                        member_names,
                        out,
                    );
                }
                if let Some(else_body) = else_body {
                    collect_java_capture_idents_stmts(
                        else_body,
                        &mut locals.clone(),
                        member_names,
                        out,
                    );
                }
            }
            StmtKind::Assign { targets, value } => {
                for target in targets {
                    collect_java_capture_idents_expr(target, locals, member_names, out);
                }
                collect_java_capture_idents_expr(value, locals, member_names, out);
            }
            StmtKind::CompoundAssign { target, value, .. } => {
                collect_java_capture_idents_expr(target, locals, member_names, out);
                collect_java_capture_idents_expr(value, locals, member_names, out);
            }
            _ => {}
        }
    }
}

fn collect_java_capture_idents_expr(
    expr: &Expression,
    locals: &std::collections::HashSet<String>,
    member_names: &JavaClassMemberNames,
    out: &mut std::collections::HashSet<String>,
) {
    match &expr.kind {
        ExprKind::Ident(name)
            if !locals.contains(name)
                && !member_names.fields.contains(name)
                && !member_names.methods.contains(name)
                && !is_java_type_or_util(name) =>
        {
            out.insert(name.clone());
        }
        ExprKind::Binary { left, right, .. } => {
            collect_java_capture_idents_expr(left, locals, member_names, out);
            collect_java_capture_idents_expr(right, locals, member_names, out);
        }
        ExprKind::Unary { expr, .. } => {
            collect_java_capture_idents_expr(expr, locals, member_names, out)
        }
        ExprKind::Member { object, .. } => {
            collect_java_capture_idents_expr(object, locals, member_names, out)
        }
        ExprKind::Call { callee, args, .. } => {
            collect_java_capture_idents_expr(callee, locals, member_names, out);
            for arg in args {
                collect_java_capture_idents_expr(&arg.value, locals, member_names, out);
            }
        }
        ExprKind::Assign { target, value } => {
            collect_java_capture_idents_expr(target, locals, member_names, out);
            collect_java_capture_idents_expr(value, locals, member_names, out);
        }
        ExprKind::Array(elems) => {
            for elem in elems {
                collect_java_capture_idents_expr(&elem.value, locals, member_names, out);
            }
        }
        _ => {}
    }
}

fn rewrite_java_anonymous_capture_refs(
    members: &mut [ClassMember],
    capture_fields: &[(String, String)],
) {
    for member in members {
        if let ClassMember::Method(method) = member {
            if let StmtKind::FunctionDecl { params, body, .. } = &mut method.kind {
                let mut locals: std::collections::HashSet<String> =
                    params.iter().map(|param| param.name.clone()).collect();
                rewrite_java_capture_refs_stmts(body, &mut locals, capture_fields);
            }
        }
    }
}

fn rewrite_java_capture_refs_stmts(
    stmts: &mut [Statement],
    locals: &mut std::collections::HashSet<String>,
    capture_fields: &[(String, String)],
) {
    for stmt in stmts {
        match &mut stmt.kind {
            StmtKind::VarDecl { declarations, .. } => {
                for decl in declarations {
                    if let Some(init) = &mut decl.init {
                        rewrite_java_capture_refs_expr(init, locals, capture_fields);
                    }
                    collect_binding_names(&decl.pattern, locals);
                }
            }
            StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
                rewrite_java_capture_refs_expr(expr, locals, capture_fields);
            }
            StmtKind::Block(body) => {
                rewrite_java_capture_refs_stmts(body, &mut locals.clone(), capture_fields)
            }
            StmtKind::Assign { targets, value } => {
                for target in targets {
                    rewrite_java_capture_refs_expr(target, locals, capture_fields);
                }
                rewrite_java_capture_refs_expr(value, locals, capture_fields);
            }
            StmtKind::CompoundAssign { target, value, .. } => {
                rewrite_java_capture_refs_expr(target, locals, capture_fields);
                rewrite_java_capture_refs_expr(value, locals, capture_fields);
            }
            _ => {}
        }
    }
}

fn rewrite_java_capture_refs_expr(
    expr: &mut Expression,
    locals: &std::collections::HashSet<String>,
    capture_fields: &[(String, String)],
) {
    match &mut expr.kind {
        ExprKind::Ident(name) if !locals.contains(name) => {
            if let Some((_, field_name)) =
                capture_fields.iter().find(|(capture, _)| capture == name)
            {
                expr.kind = ExprKind::Member {
                    object: Box::new(Expression::new(ExprKind::This)),
                    field: field_name.clone(),
                    null_safe: false,
                };
            }
        }
        ExprKind::Binary { left, right, .. } => {
            rewrite_java_capture_refs_expr(left, locals, capture_fields);
            rewrite_java_capture_refs_expr(right, locals, capture_fields);
        }
        ExprKind::Unary { expr, .. } => {
            rewrite_java_capture_refs_expr(expr, locals, capture_fields)
        }
        ExprKind::Member { object, .. } => {
            rewrite_java_capture_refs_expr(object, locals, capture_fields)
        }
        ExprKind::Call { callee, args, .. } => {
            rewrite_java_capture_refs_expr(callee, locals, capture_fields);
            for arg in args {
                rewrite_java_capture_refs_expr(&mut arg.value, locals, capture_fields);
            }
        }
        ExprKind::Assign { target, value } => {
            rewrite_java_capture_refs_expr(target, locals, capture_fields);
            rewrite_java_capture_refs_expr(value, locals, capture_fields);
        }
        ExprKind::Array(elems) => {
            for elem in elems {
                rewrite_java_capture_refs_expr(&mut elem.value, locals, capture_fields);
            }
        }
        _ => {}
    }
}

fn normalize_java_anonymous_class_members(
    members: &mut [ClassMember],
    class_members: &std::collections::HashMap<String, JavaClassMemberNames>,
) {
    for member in members {
        match member {
            ClassMember::Field { init, .. } => {
                if let Some(init) = init {
                    normalize_java_anonymous_class_expr(init, class_members);
                }
            }
            ClassMember::Constructor { body, .. } => {
                normalize_java_anonymous_class_tree(body, class_members);
            }
            ClassMember::Method(method) => {
                if let StmtKind::FunctionDecl { body, .. } = &mut method.kind {
                    normalize_java_anonymous_class_tree(body, class_members);
                }
            }
            ClassMember::NestedType(nested) => {
                normalize_java_anonymous_class_stmt(nested, class_members);
            }
            _ => {}
        }
    }
}

fn normalize_java_anonymous_class_stmt(
    stmt: &mut Statement,
    class_members: &std::collections::HashMap<String, JavaClassMemberNames>,
) {
    match &mut stmt.kind {
        StmtKind::Expr(expr)
        | StmtKind::Return(Some(expr))
        | StmtKind::Throw {
            expr: Some(expr), ..
        } => {
            normalize_java_anonymous_class_expr(expr, class_members);
        }
        StmtKind::Block(stmts) | StmtKind::NamespaceDecl { body: stmts, .. } => {
            normalize_java_anonymous_class_tree(stmts, class_members);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    normalize_java_anonymous_class_expr(init, class_members);
                }
            }
        }
        StmtKind::FunctionDecl { body, .. } => {
            normalize_java_anonymous_class_tree(body, class_members)
        }
        StmtKind::ClassDecl {
            members,
            decorators,
            ..
        }
        | StmtKind::StructDecl {
            members,
            decorators,
            ..
        }
        | StmtKind::EnumDecl {
            body_members: members,
            decorators,
            ..
        } => {
            for decorator in decorators {
                normalize_java_anonymous_class_expr(decorator, class_members);
            }
            normalize_java_anonymous_class_members(members, class_members);
        }
        StmtKind::InterfaceDecl { decorators, .. } => {
            for decorator in decorators {
                normalize_java_anonymous_class_expr(decorator, class_members);
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            normalize_java_anonymous_class_expr(cond, class_members);
            normalize_java_anonymous_class_tree(then_body, class_members);
            for (elif_cond, elif_body) in elifs {
                normalize_java_anonymous_class_expr(elif_cond, class_members);
                normalize_java_anonymous_class_tree(elif_body, class_members);
            }
            if let Some(else_body) = else_body {
                normalize_java_anonymous_class_tree(else_body, class_members);
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                normalize_java_anonymous_class_stmt(init, class_members);
            }
            if let Some(cond) = cond {
                normalize_java_anonymous_class_expr(cond, class_members);
            }
            if let Some(update) = update {
                normalize_java_anonymous_class_expr(update, class_members);
            }
            normalize_java_anonymous_class_tree(body, class_members);
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            normalize_java_anonymous_class_expr(iter, class_members);
            normalize_java_anonymous_class_tree(body, class_members);
            if let Some(else_body) = else_body {
                normalize_java_anonymous_class_tree(else_body, class_members);
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            normalize_java_anonymous_class_expr(cond, class_members);
            normalize_java_anonymous_class_tree(body, class_members);
            if let Some(else_body) = else_body {
                normalize_java_anonymous_class_tree(else_body, class_members);
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            normalize_java_anonymous_class_tree(body, class_members);
            normalize_java_anonymous_class_expr(cond, class_members);
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            normalize_java_anonymous_class_expr(expr, class_members);
            for case in cases {
                for condition in &mut case.conditions {
                    normalize_java_anonymous_case_condition(condition, class_members);
                }
                normalize_java_anonymous_class_tree(&mut case.body, class_members);
            }
            if let Some(default) = default {
                normalize_java_anonymous_class_tree(default, class_members);
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            normalize_java_anonymous_class_tree(body, class_members);
            for catch in catches {
                normalize_java_anonymous_class_tree(&mut catch.body, class_members);
            }
            if let Some(else_body) = else_body {
                normalize_java_anonymous_class_tree(else_body, class_members);
            }
            if let Some(finally) = finally {
                normalize_java_anonymous_class_tree(finally, class_members);
            }
        }
        StmtKind::With { items, body, .. } => {
            for item in items {
                normalize_java_anonymous_class_expr(&mut item.expr, class_members);
            }
            normalize_java_anonymous_class_tree(body, class_members);
        }
        StmtKind::Using { resource, body, .. } => {
            normalize_java_anonymous_class_expr(resource, class_members);
            normalize_java_anonymous_class_tree(body, class_members);
        }
        StmtKind::Lock { expr, body } => {
            normalize_java_anonymous_class_expr(expr, class_members);
            normalize_java_anonymous_class_tree(body, class_members);
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                normalize_java_anonymous_class_expr(target, class_members);
            }
            normalize_java_anonymous_class_expr(value, class_members);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            normalize_java_anonymous_class_expr(target, class_members);
            normalize_java_anonymous_class_expr(value, class_members);
        }
        StmtKind::AddHandler {
            control, handler, ..
        } => {
            normalize_java_anonymous_class_expr(control, class_members);
            normalize_java_anonymous_class_expr(handler, class_members);
        }
        StmtKind::RemoveHandler {
            control, handler, ..
        } => {
            normalize_java_anonymous_class_expr(control, class_members);
            normalize_java_anonymous_class_expr(handler, class_members);
        }
        StmtKind::RaiseEvent { args, .. }
        | StmtKind::PrintFile { items: args, .. }
        | StmtKind::WriteFile { items: args, .. }
        | StmtKind::Echo(args)
        | StmtKind::Delete(args) => {
            for arg in args {
                normalize_java_anonymous_class_expr(arg, class_members);
            }
        }
        StmtKind::OpenFile {
            path, file_number, ..
        } => {
            normalize_java_anonymous_class_expr(path, class_members);
            normalize_java_anonymous_class_expr(file_number, class_members);
        }
        StmtKind::CloseFile(Some(expr)) | StmtKind::Assert { test: expr, .. } => {
            normalize_java_anonymous_class_expr(expr, class_members);
        }
        StmtKind::InputFile {
            file_number,
            variables,
        } => {
            normalize_java_anonymous_class_expr(file_number, class_members);
            for variable in variables {
                normalize_java_anonymous_class_expr(variable, class_members);
            }
        }
        StmtKind::LineInput { file_number, .. } => {
            normalize_java_anonymous_class_expr(file_number, class_members);
        }
        StmtKind::StartFile {
            file_number,
            key_value,
            ..
        } => {
            normalize_java_anonymous_class_expr(file_number, class_members);
            normalize_java_anonymous_class_expr(key_value, class_members);
        }
        StmtKind::InputRecordFile {
            file_number,
            key_value,
            ..
        } => {
            normalize_java_anonymous_class_expr(file_number, class_members);
            if let Some(key_value) = key_value {
                normalize_java_anonymous_class_expr(key_value, class_members);
            }
        }
        StmtKind::RewriteRecordFile {
            file_number, items, ..
        } => {
            normalize_java_anonymous_class_expr(file_number, class_members);
            for item in items {
                normalize_java_anonymous_class_expr(item, class_members);
            }
        }
        StmtKind::Export {
            declaration,
            default,
            ..
        } => {
            if let Some(declaration) = declaration {
                normalize_java_anonymous_class_stmt(declaration, class_members);
            }
            if let Some(default) = default {
                normalize_java_anonymous_class_expr(default, class_members);
            }
        }
        StmtKind::Labeled { body, .. } => normalize_java_anonymous_class_stmt(body, class_members),
        StmtKind::MatchStatement { subject, cases } => {
            normalize_java_anonymous_class_expr(subject, class_members);
            for case in cases {
                if let Some(guard) = &mut case.guard {
                    normalize_java_anonymous_class_expr(guard, class_members);
                }
                normalize_java_anonymous_class_tree(&mut case.body, class_members);
            }
        }
        _ => {}
    }
}

fn normalize_java_anonymous_class_expr(
    expr: &mut Expression,
    class_members: &std::collections::HashMap<String, JavaClassMemberNames>,
) {
    match &mut expr.kind {
        ExprKind::Binary { left, right, .. } => {
            normalize_java_anonymous_class_expr(left, class_members);
            normalize_java_anonymous_class_expr(right, class_members);
        }
        ExprKind::Unary { expr: inner, .. }
        | ExprKind::Spread(inner)
        | ExprKind::Await(inner)
        | ExprKind::YieldFrom(inner)
        | ExprKind::Void(inner)
        | ExprKind::Delete(inner)
        | ExprKind::TypeOf(inner)
        | ExprKind::RefLoad(inner) => {
            normalize_java_anonymous_class_expr(inner, class_members);
        }
        ExprKind::Yield(Some(inner)) => normalize_java_anonymous_class_expr(inner, class_members),
        ExprKind::Ternary { cond, then, else_ } => {
            normalize_java_anonymous_class_expr(cond, class_members);
            normalize_java_anonymous_class_expr(then, class_members);
            normalize_java_anonymous_class_expr(else_, class_members);
        }
        ExprKind::Member { object, .. } => {
            normalize_java_anonymous_class_expr(object, class_members)
        }
        ExprKind::Index { object, index, .. } => {
            normalize_java_anonymous_class_expr(object, class_members);
            normalize_java_anonymous_class_expr(index, class_members);
        }
        ExprKind::Call { callee, args, .. } => {
            normalize_java_anonymous_class_expr(callee, class_members);
            for arg in args {
                normalize_java_anonymous_class_expr(&mut arg.value, class_members);
            }
        }
        ExprKind::New { class, args } => {
            normalize_java_anonymous_class_expr(class, class_members);
            for arg in args {
                normalize_java_anonymous_class_expr(&mut arg.value, class_members);
            }
        }
        ExprKind::Assign { target, value } => {
            normalize_java_anonymous_class_expr(target, class_members);
            normalize_java_anonymous_class_expr(value, class_members);
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(inner) => normalize_java_anonymous_class_expr(inner, class_members),
            LambdaBody::Block(stmts) => normalize_java_anonymous_class_tree(stmts, class_members),
        },
        ExprKind::Array(elems) => {
            for elem in elems {
                normalize_java_anonymous_class_expr(&mut elem.value, class_members);
            }
        }
        ExprKind::Tuple(elems) | ExprKind::Set(elems) | ExprKind::Sequence(elems) => {
            for elem in elems {
                normalize_java_anonymous_class_expr(elem, class_members);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                normalize_java_anonymous_object_property(prop, class_members);
            }
        }
        ExprKind::Interpolation(parts) => {
            for part in parts {
                if let InterpolPart::Expr(inner) = part {
                    normalize_java_anonymous_class_expr(inner, class_members);
                }
            }
        }
        ExprKind::IsType { expr: inner, .. }
        | ExprKind::Cast { expr: inner, .. }
        | ExprKind::NullCoalesce { left: inner, .. } => {
            normalize_java_anonymous_class_expr(inner, class_members);
            if let ExprKind::NullCoalesce { right, .. } = &mut expr.kind {
                normalize_java_anonymous_class_expr(right, class_members);
            }
        }
        ExprKind::Comprehension {
            element,
            generators,
            ..
        } => {
            normalize_java_anonymous_class_expr(element, class_members);
            for generator in generators {
                normalize_java_anonymous_class_expr(&mut generator.target, class_members);
                normalize_java_anonymous_class_expr(&mut generator.iter, class_members);
                for condition in &mut generator.conditions {
                    normalize_java_anonymous_class_expr(condition, class_members);
                }
            }
        }
        ExprKind::Slice { lower, upper, step } => {
            for part in [lower, upper, step] {
                if let Some(part) = part {
                    normalize_java_anonymous_class_expr(part, class_members);
                }
            }
        }
        ExprKind::Walrus { target, value } => {
            normalize_java_anonymous_class_expr(target, class_members);
            normalize_java_anonymous_class_expr(value, class_members);
        }
        ExprKind::ClassExpr {
            parent, members, ..
        } => {
            if let Some(parent) = parent {
                normalize_java_anonymous_class_expr(parent, class_members);
            }
            let mut names = JavaClassMemberNames::from_members(members);
            if let Some(parent_name) = parent.as_deref().and_then(java_parent_expr_name) {
                if let Some(parent_names) = class_members.get(parent_name) {
                    names.fields.extend(parent_names.fields.iter().cloned());
                    names.methods.extend(parent_names.methods.iter().cloned());
                    names
                        .static_methods
                        .extend(parent_names.static_methods.iter().cloned());
                }
            }
            normalize_java_class_members(members, "", &names, class_members);
            normalize_java_anonymous_class_members(members, class_members);
        }
        ExprKind::FunctionExpr(func) => normalize_java_anonymous_class_stmt(func, class_members),
        ExprKind::Range { start, end, .. } => {
            normalize_java_anonymous_class_expr(start, class_members);
            normalize_java_anonymous_class_expr(end, class_members);
        }
        ExprKind::StaticAccess { class, member } => {
            normalize_java_anonymous_class_expr(class, class_members);
            normalize_java_anonymous_class_expr(member, class_members);
        }
        ExprKind::Match { subject, arms } => {
            normalize_java_anonymous_class_expr(subject, class_members);
            for arm in arms {
                if let Some(conditions) = &mut arm.conditions {
                    for condition in conditions {
                        normalize_java_anonymous_class_expr(condition, class_members);
                    }
                }
                normalize_java_anonymous_class_expr(&mut arm.body, class_members);
            }
        }
        _ => {}
    }
}

fn normalize_java_anonymous_case_condition(
    condition: &mut CaseCondition,
    class_members: &std::collections::HashMap<String, JavaClassMemberNames>,
) {
    match condition {
        CaseCondition::Value(expr) | CaseCondition::Comparison { expr, .. } => {
            normalize_java_anonymous_class_expr(expr, class_members);
        }
        CaseCondition::Range { from, to } => {
            normalize_java_anonymous_class_expr(from, class_members);
            normalize_java_anonymous_class_expr(to, class_members);
        }
    }
}

fn normalize_java_anonymous_object_property(
    prop: &mut ObjectProperty,
    class_members: &std::collections::HashMap<String, JavaClassMemberNames>,
) {
    match prop {
        ObjectProperty::KeyValue { key, value } | ObjectProperty::Computed { key, value } => {
            normalize_java_anonymous_class_expr(key, class_members);
            normalize_java_anonymous_class_expr(value, class_members);
        }
        ObjectProperty::Spread(expr) => normalize_java_anonymous_class_expr(expr, class_members),
        ObjectProperty::Method { value, .. } | ObjectProperty::Accessor { value, .. } => {
            normalize_java_anonymous_class_stmt(value, class_members);
        }
        ObjectProperty::Shorthand(_) => {}
    }
}

fn java_parent_expr_name(expr: &Expression) -> Option<&str> {
    match &expr.kind {
        ExprKind::Ident(name) => Some(name.as_str()),
        ExprKind::Member { field, .. } => Some(field.as_str()),
        _ => None,
    }
}

#[derive(Clone, Default)]
struct JavaClassMemberNames {
    fields: std::collections::HashSet<String>,
    field_types: HashMap<String, String>,
    methods: std::collections::HashSet<String>,
    static_methods: std::collections::HashSet<String>,
    default_methods: Vec<Box<Statement>>,
    instance_overloads: HashMap<String, Vec<JavaOverloadTarget>>,
    static_overloads: HashMap<String, Vec<JavaOverloadTarget>>,
}

#[derive(Clone)]
struct JavaOverloadTarget {
    mangled_name: String,
    param_types: Vec<String>,
    return_type: Option<String>,
}

impl JavaClassMemberNames {
    fn from_members(members: &[ClassMember]) -> Self {
        let fields = members
            .iter()
            .filter_map(|member| match member {
                ClassMember::Field {
                    name, modifiers, ..
                } if !modifiers.is_static => Some(name.clone()),
                _ => None,
            })
            .collect();
        let field_types = members
            .iter()
            .filter_map(|member| match member {
                ClassMember::Field {
                    name, type_hint, ..
                } => type_hint.as_ref().map(|t| (name.clone(), t.clone())),
                _ => None,
            })
            .collect();
        let methods = members
            .iter()
            .filter_map(|member| match member {
                ClassMember::Method(func) => match &func.kind {
                    StmtKind::FunctionDecl {
                        name, modifiers, ..
                    } if !modifiers.is_static => Some(name.clone()),
                    _ => None,
                },
                _ => None,
            })
            .collect();
        let static_methods = members
            .iter()
            .filter_map(|member| match member {
                ClassMember::Method(func) => match &func.kind {
                    StmtKind::FunctionDecl {
                        name, modifiers, ..
                    } if modifiers.is_static => Some(name.clone()),
                    _ => None,
                },
                _ => None,
            })
            .collect();
        let default_methods = members
            .iter()
            .filter_map(|member| match member {
                ClassMember::Method(func) => match &func.kind {
                    StmtKind::FunctionDecl {
                        body, modifiers, ..
                    } if !modifiers.is_static && !body.is_empty() => Some(func.clone()),
                    _ => None,
                },
                _ => None,
            })
            .collect();
        let instance_overloads = java_collect_overload_targets(members, false);
        let static_overloads = java_collect_overload_targets(members, true);
        Self {
            fields,
            field_types,
            methods,
            static_methods,
            default_methods,
            instance_overloads,
            static_overloads,
        }
    }
}

fn normalize_java_class_members(
    members: &mut [ClassMember],
    class_name: &str,
    names: &JavaClassMemberNames,
    class_members: &std::collections::HashMap<String, JavaClassMemberNames>,
) {
    if names.fields.is_empty() && names.methods.is_empty() && names.static_methods.is_empty() {
        return;
    }

    let current_class = if class_name.is_empty() {
        None
    } else {
        Some(class_name)
    };
    normalize_java_overload_names(members);

    for member in members {
        match member {
            ClassMember::Constructor { params, body, .. } => {
                let mut locals = params.iter().map(|p| p.name.clone()).collect();
                let mut local_types: HashMap<String, String> = params
                    .iter()
                    .filter_map(|p| p.type_hint.as_ref().map(|t| (p.name.clone(), t.clone())))
                    .collect();
                local_types.extend(
                    names
                        .field_types
                        .iter()
                        .map(|(name, ty)| (name.clone(), ty.clone())),
                );
                normalize_java_stmts(
                    body,
                    &names.fields,
                    &names.methods,
                    &names.static_methods,
                    &names.static_overloads,
                    class_members,
                    current_class,
                    &mut locals,
                    &mut local_types,
                );
            }
            ClassMember::Method(func) => {
                if let StmtKind::FunctionDecl {
                    params,
                    body,
                    modifiers,
                    ..
                } = &mut func.kind
                {
                    let mut locals = params.iter().map(|p| p.name.clone()).collect();
                    let mut local_types: HashMap<String, String> = params
                        .iter()
                        .filter_map(|p| p.type_hint.as_ref().map(|t| (p.name.clone(), t.clone())))
                        .collect();
                    local_types.extend(
                        names
                            .field_types
                            .iter()
                            .map(|(name, ty)| (name.clone(), ty.clone())),
                    );
                    if modifiers.is_static {
                        let empty = std::collections::HashSet::new();
                        normalize_java_stmts(
                            body,
                            &empty,
                            &empty,
                            &names.static_methods,
                            &names.static_overloads,
                            class_members,
                            current_class,
                            &mut locals,
                            &mut local_types,
                        );
                    } else {
                        normalize_java_stmts(
                            body,
                            &names.fields,
                            &names.methods,
                            &names.static_methods,
                            &names.static_overloads,
                            class_members,
                            current_class,
                            &mut locals,
                            &mut local_types,
                        );
                    }
                }
            }
            _ => {}
        }
    }
}

fn java_collect_overload_targets(
    members: &[ClassMember],
    want_static: bool,
) -> HashMap<String, Vec<JavaOverloadTarget>> {
    let mut counts: HashMap<String, usize> = HashMap::new();
    for member in members {
        let ClassMember::Method(func) = member else {
            continue;
        };
        let StmtKind::FunctionDecl { name, .. } = &func.kind else {
            continue;
        };
        *counts.entry(name.clone()).or_default() += 1;
    }

    let mut overloads: HashMap<String, Vec<JavaOverloadTarget>> = HashMap::new();
    for member in members {
        let ClassMember::Method(func) = member else {
            continue;
        };
        let StmtKind::FunctionDecl {
            name,
            params,
            modifiers,
            return_type,
            ..
        } = &func.kind
        else {
            continue;
        };
        if modifiers.is_static != want_static || counts.get(name).copied().unwrap_or(0) < 2 {
            continue;
        }
        let param_types: Vec<String> = params
            .iter()
            .map(|param| java_overload_type_key(param.type_hint.as_deref().unwrap_or("Object")))
            .collect();
        overloads
            .entry(name.clone())
            .or_default()
            .push(JavaOverloadTarget {
                mangled_name: java_overload_mangled_name(name, &param_types),
                param_types,
                return_type: return_type.clone(),
            });
    }
    overloads
}

fn normalize_java_overload_names(members: &mut [ClassMember]) {
    let mut counts: HashMap<String, usize> = HashMap::new();
    for member in members.iter() {
        let ClassMember::Method(func) = member else {
            continue;
        };
        let StmtKind::FunctionDecl { name, .. } = &func.kind else {
            continue;
        };
        *counts.entry(name.clone()).or_default() += 1;
    }

    for member in members.iter_mut() {
        let ClassMember::Method(func) = member else {
            continue;
        };
        let StmtKind::FunctionDecl { name, params, .. } = &mut func.kind else {
            continue;
        };
        if counts.get(name).copied().unwrap_or(0) < 2 {
            continue;
        }
        let original = name.clone();
        let param_types: Vec<String> = params
            .iter()
            .map(|param| java_overload_type_key(param.type_hint.as_deref().unwrap_or("Object")))
            .collect();
        *name = java_overload_mangled_name(&original, &param_types);
    }
}

fn java_overload_mangled_name(name: &str, param_types: &[String]) -> String {
    format!(
        "{}__java_{}",
        name,
        if param_types.is_empty() {
            "void".to_string()
        } else {
            param_types.join("_")
        }
    )
}

fn java_overload_type_key(type_hint: &str) -> String {
    let trimmed = type_hint.trim().trim_end_matches("[]").trim();
    let simple = java_type_simple_name(trimmed);
    match simple {
        "Integer" => "int".to_string(),
        "Double" => "double".to_string(),
        "Float" => "float".to_string(),
        "Long" => "long".to_string(),
        "Short" => "short".to_string(),
        "Byte" => "byte".to_string(),
        "Boolean" => "boolean".to_string(),
        "Character" => "char".to_string(),
        "String" => "string".to_string(),
        other => other
            .chars()
            .map(|ch| if ch.is_ascii_alphanumeric() { ch } else { '_' })
            .collect::<String>()
            .to_ascii_lowercase(),
    }
}

fn java_expr_overload_type(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(_)) => Some("int".to_string()),
        ExprKind::Lit(Literal::Float(_)) => Some("double".to_string()),
        ExprKind::Lit(Literal::Str(_)) => Some("string".to_string()),
        ExprKind::Lit(Literal::Bool(_)) => Some("boolean".to_string()),
        ExprKind::Lit(Literal::Char(_)) => Some("char".to_string()),
        ExprKind::Cast { type_name, .. } => Some(java_overload_type_key(type_name)),
        ExprKind::Unary {
            op: UnaryOp::Neg | UnaryOp::Pos,
            expr,
        } => java_expr_overload_type(expr),
        _ => None,
    }
}

fn select_java_overload_target<'a>(
    method: &str,
    args: &[Argument],
    overloads: &'a HashMap<String, Vec<JavaOverloadTarget>>,
) -> Option<&'a JavaOverloadTarget> {
    let targets = overloads.get(method)?;
    let arg_types: Option<Vec<String>> = args
        .iter()
        .map(|arg| java_expr_overload_type(&arg.value))
        .collect();
    let arg_types = arg_types?;
    if let Some(target) = targets
        .iter()
        .find(|target| target.param_types == arg_types)
    {
        return Some(target);
    }
    targets
        .iter()
        .find(|target| java_overload_args_match_with_char_fallback(&target.param_types, args))
}

fn java_overload_args_match_with_char_fallback(param_types: &[String], args: &[Argument]) -> bool {
    param_types.len() == args.len()
        && param_types
            .iter()
            .zip(args.iter())
            .all(|(param_type, arg)| match java_expr_overload_type(&arg.value) {
                Some(arg_type) if &arg_type == param_type => true,
                Some(arg_type) if arg_type == "string" && param_type == "char" => {
                    matches!(&arg.value.kind, ExprKind::Lit(Literal::Str(s)) if s.chars().count() == 1)
                }
                _ => false,
            })
}

fn java_overload_return_cast_type(return_type: Option<&str>) -> Option<String> {
    match return_type.map(java_type_simple_name) {
        Some("double") | Some("Double") => Some("double".to_string()),
        Some("float") | Some("Float") => Some("float".to_string()),
        _ => None,
    }
}

fn java_overload_receiver_type(
    object: &Expression,
    local_types: &HashMap<String, String>,
) -> Option<String> {
    match &object.kind {
        ExprKind::Ident(name) => local_types.get(name).cloned(),
        ExprKind::This => None,
        ExprKind::New { class, .. } => java_expr_dotted_name(class),
        _ => None,
    }
}

fn java_static_field_receiver(
    object: &Expression,
    current_class: Option<&str>,
) -> Option<Expression> {
    let class_name = current_class?;
    let ExprKind::Ident(field_name) = &object.kind else {
        return None;
    };
    let key = format!("{class_name}.{field_name}");
    let is_static = JAVA_STATIC_FIELD_VARS.with(|vars| vars.borrow().contains(&key));
    if !is_static {
        return None;
    }
    Some(Expression::new(ExprKind::Member {
        object: Box::new(Expression::ident(class_name)),
        field: field_name.clone(),
        null_safe: false,
    }))
}

fn java_static_field_type(object: &Expression) -> Option<String> {
    let ExprKind::Member {
        object: class,
        field,
        ..
    } = &object.kind
    else {
        return None;
    };
    let ExprKind::Ident(class_name) = &class.kind else {
        return None;
    };
    let key = format!("{class_name}.{field}");
    JAVA_STATIC_FIELD_TYPES.with(|types| types.borrow().get(&key).cloned())
}

fn java_current_static_field_type(field_name: &str) -> Option<(String, String)> {
    let class_name = JAVA_CURRENT_CLASS_STACK.with(|stack| stack.borrow().last().cloned())?;
    let key = format!("{class_name}.{field_name}");
    JAVA_STATIC_FIELD_TYPES
        .with(|types| types.borrow().get(&key).cloned())
        .map(|ty| (class_name, ty))
}

fn java_receiver_type(
    object: &Expression,
    local_types: &HashMap<String, String>,
) -> Option<String> {
    java_overload_receiver_type(object, local_types).or_else(|| java_static_field_type(object))
}

fn java_lookup_class_members<'a>(
    class_members: &'a std::collections::HashMap<String, JavaClassMemberNames>,
    type_name: &str,
) -> Option<&'a JavaClassMemberNames> {
    let simple = java_type_simple_name(type_name);
    class_members
        .get(type_name)
        .or_else(|| class_members.get(simple))
}

fn java_expr_reads_any_local(expr: &Expression, locals: &HashSet<String>) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => locals.contains(name),
        ExprKind::Lambda { params, body, .. } => {
            let mut nested = locals.clone();
            for param in params {
                nested.remove(&param.name);
            }
            match body {
                LambdaBody::Expr(inner) => java_expr_reads_any_local(inner, &nested),
                LambdaBody::Block(stmts) => stmts
                    .iter()
                    .any(|stmt| java_stmt_reads_any_local(stmt, &nested)),
            }
        }
        ExprKind::Call { callee, args, .. } => {
            java_expr_reads_any_local(callee, locals)
                || args
                    .iter()
                    .any(|arg| java_expr_reads_any_local(&arg.value, locals))
        }
        ExprKind::Member { object, .. } => java_expr_reads_any_local(object, locals),
        ExprKind::Index { object, index, .. } => {
            java_expr_reads_any_local(object, locals) || java_expr_reads_any_local(index, locals)
        }
        ExprKind::Unary { expr, .. } => java_expr_reads_any_local(expr, locals),
        ExprKind::Binary { left, right, .. } => {
            java_expr_reads_any_local(left, locals) || java_expr_reads_any_local(right, locals)
        }
        ExprKind::Assign { target, value, .. } => {
            java_expr_reads_any_local(target, locals) || java_expr_reads_any_local(value, locals)
        }
        ExprKind::Ternary { cond, then, else_ } => {
            java_expr_reads_any_local(cond, locals)
                || java_expr_reads_any_local(then, locals)
                || java_expr_reads_any_local(else_, locals)
        }
        ExprKind::New { class, args, .. } => {
            java_expr_reads_any_local(class, locals)
                || args
                    .iter()
                    .any(|arg| java_expr_reads_any_local(&arg.value, locals))
        }
        ExprKind::Array(items) => items.iter().any(|item| {
            item.key
                .as_ref()
                .is_some_and(|key| java_expr_reads_any_local(key, locals))
                || java_expr_reads_any_local(&item.value, locals)
        }),
        ExprKind::Object(props) => props.iter().any(|prop| match prop {
            ObjectProperty::KeyValue { key, value } => {
                java_expr_reads_any_local(key, locals) || java_expr_reads_any_local(value, locals)
            }
            ObjectProperty::Spread(value) => java_expr_reads_any_local(value, locals),
            ObjectProperty::Shorthand(name) => locals.contains(name),
            _ => true,
        }),
        ExprKind::Sequence(exprs) => exprs
            .iter()
            .any(|expr| java_expr_reads_any_local(expr, locals)),
        _ => false,
    }
}

fn java_stmt_reads_any_local(stmt: &Statement, locals: &HashSet<String>) -> bool {
    match &stmt.kind {
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            java_expr_reads_any_local(expr, locals)
        }
        StmtKind::VarDecl { declarations, .. } => declarations.iter().any(|decl| {
            decl.init
                .as_ref()
                .is_some_and(|init| java_expr_reads_any_local(init, locals))
        }),
        StmtKind::Block(stmts) => stmts
            .iter()
            .any(|stmt| java_stmt_reads_any_local(stmt, locals)),
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            java_expr_reads_any_local(cond, locals)
                || then_body
                    .iter()
                    .any(|stmt| java_stmt_reads_any_local(stmt, locals))
                || elifs.iter().any(|(cond, body)| {
                    java_expr_reads_any_local(cond, locals)
                        || body
                            .iter()
                            .any(|stmt| java_stmt_reads_any_local(stmt, locals))
                })
                || else_body.as_ref().is_some_and(|body| {
                    body.iter()
                        .any(|stmt| java_stmt_reads_any_local(stmt, locals))
                })
        }
        StmtKind::While { cond, body, .. } => {
            java_expr_reads_any_local(cond, locals)
                || body
                    .iter()
                    .any(|stmt| java_stmt_reads_any_local(stmt, locals))
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            init.as_ref()
                .is_some_and(|stmt| java_stmt_reads_any_local(stmt, locals))
                || cond
                    .as_ref()
                    .is_some_and(|expr| java_expr_reads_any_local(expr, locals))
                || update
                    .as_ref()
                    .is_some_and(|expr| java_expr_reads_any_local(expr, locals))
                || body
                    .iter()
                    .any(|stmt| java_stmt_reads_any_local(stmt, locals))
        }
        StmtKind::ForIn { iter, body, .. } => {
            java_expr_reads_any_local(iter, locals)
                || body
                    .iter()
                    .any(|stmt| java_stmt_reads_any_local(stmt, locals))
        }
        StmtKind::Try {
            body,
            catches,
            finally,
            ..
        } => {
            body.iter()
                .any(|stmt| java_stmt_reads_any_local(stmt, locals))
                || catches.iter().any(|catch| {
                    catch
                        .body
                        .iter()
                        .any(|stmt| java_stmt_reads_any_local(stmt, locals))
                })
                || finally.as_ref().is_some_and(|body| {
                    body.iter()
                        .any(|stmt| java_stmt_reads_any_local(stmt, locals))
                })
        }
        StmtKind::Return(None) | StmtKind::Break(_) | StmtKind::Continue(_) => false,
        _ => false,
    }
}

fn java_thread_target_is_unsafe(target: &Expression, locals: &HashSet<String>) -> bool {
    match &target.kind {
        ExprKind::Ident(name) => {
            JAVA_RUNNABLE_UNSAFE_TARGETS.with(|targets| targets.borrow().contains(name))
        }
        ExprKind::Lambda { .. } => java_expr_reads_any_local(target, locals),
        _ => false,
    }
}

fn java_resolve_runnable_target(target: Expression) -> Expression {
    if let ExprKind::Ident(name) = &target.kind {
        if let Some(resolved) =
            JAVA_RUNNABLE_TARGETS.with(|targets| targets.borrow().get(name).cloned())
        {
            return resolved;
        }
    }
    target
}

fn java_rewrite_spawned_thread_sleep_expr(expr: &mut Expression) {
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            java_rewrite_spawned_thread_sleep_expr(callee);
            for arg in &mut *args {
                java_rewrite_spawned_thread_sleep_expr(&mut arg.value);
            }
            if matches!(&callee.kind, ExprKind::Ident(name) if name == "__j_thread_sleep") {
                *callee = Box::new(Expression::ident("__java_thread_sleep"));
            }
            if let ExprKind::Member { object, field, .. } = &callee.kind {
                if field == "put" && args.len() == 1 {
                    let mut new_args = Vec::with_capacity(2);
                    new_args.push(Argument::positional((**object).clone()));
                    new_args.extend(args.iter().cloned());
                    *expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__java_blocking_queue_put")),
                        args: new_args,
                        optional: false,
                    });
                }
            }
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(inner) => java_rewrite_spawned_thread_sleep_expr(inner),
            LambdaBody::Block(stmts) => {
                for stmt in stmts {
                    java_rewrite_spawned_thread_sleep_stmt(stmt);
                }
            }
        },
        ExprKind::Member { object, .. } => java_rewrite_spawned_thread_sleep_expr(object),
        ExprKind::Index { object, index, .. } => {
            java_rewrite_spawned_thread_sleep_expr(object);
            java_rewrite_spawned_thread_sleep_expr(index);
        }
        ExprKind::Unary { expr, .. } => java_rewrite_spawned_thread_sleep_expr(expr),
        ExprKind::Binary { left, right, .. } => {
            java_rewrite_spawned_thread_sleep_expr(left);
            java_rewrite_spawned_thread_sleep_expr(right);
        }
        ExprKind::Assign { target, value } => {
            java_rewrite_spawned_thread_sleep_expr(target);
            java_rewrite_spawned_thread_sleep_expr(value);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            java_rewrite_spawned_thread_sleep_expr(cond);
            java_rewrite_spawned_thread_sleep_expr(then);
            java_rewrite_spawned_thread_sleep_expr(else_);
        }
        ExprKind::New { class, args, .. } => {
            java_rewrite_spawned_thread_sleep_expr(class);
            for arg in args {
                java_rewrite_spawned_thread_sleep_expr(&mut arg.value);
            }
        }
        ExprKind::Array(items) => {
            for item in items {
                if let Some(key) = &mut item.key {
                    java_rewrite_spawned_thread_sleep_expr(key);
                }
                java_rewrite_spawned_thread_sleep_expr(&mut item.value);
            }
        }
        ExprKind::Sequence(exprs) => {
            for expr in exprs {
                java_rewrite_spawned_thread_sleep_expr(expr);
            }
        }
        _ => {}
    }
}

fn java_rewrite_spawned_thread_sleep_stmt(stmt: &mut Statement) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            java_rewrite_spawned_thread_sleep_expr(expr);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    java_rewrite_spawned_thread_sleep_expr(init);
                }
            }
        }
        StmtKind::Block(stmts) => {
            for stmt in stmts {
                java_rewrite_spawned_thread_sleep_stmt(stmt);
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            java_rewrite_spawned_thread_sleep_expr(cond);
            for stmt in then_body {
                java_rewrite_spawned_thread_sleep_stmt(stmt);
            }
            for (cond, body) in elifs {
                java_rewrite_spawned_thread_sleep_expr(cond);
                for stmt in body {
                    java_rewrite_spawned_thread_sleep_stmt(stmt);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    java_rewrite_spawned_thread_sleep_stmt(stmt);
                }
            }
        }
        StmtKind::While { cond, body, .. } => {
            java_rewrite_spawned_thread_sleep_expr(cond);
            for stmt in body {
                java_rewrite_spawned_thread_sleep_stmt(stmt);
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                java_rewrite_spawned_thread_sleep_stmt(init);
            }
            if let Some(cond) = cond {
                java_rewrite_spawned_thread_sleep_expr(cond);
            }
            if let Some(update) = update {
                java_rewrite_spawned_thread_sleep_expr(update);
            }
            for stmt in body {
                java_rewrite_spawned_thread_sleep_stmt(stmt);
            }
        }
        StmtKind::ForIn { iter, body, .. } => {
            java_rewrite_spawned_thread_sleep_expr(iter);
            for stmt in body {
                java_rewrite_spawned_thread_sleep_stmt(stmt);
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
            ..
        } => {
            for stmt in body {
                java_rewrite_spawned_thread_sleep_stmt(stmt);
            }
            for catch in catches {
                for stmt in &mut catch.body {
                    java_rewrite_spawned_thread_sleep_stmt(stmt);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    java_rewrite_spawned_thread_sleep_stmt(stmt);
                }
            }
            if let Some(body) = finally {
                for stmt in body {
                    java_rewrite_spawned_thread_sleep_stmt(stmt);
                }
            }
        }
        _ => {}
    }
}

fn normalize_java_stmts(
    stmts: &mut [Statement],
    fields: &std::collections::HashSet<String>,
    methods: &std::collections::HashSet<String>,
    static_methods: &std::collections::HashSet<String>,
    static_overloads: &HashMap<String, Vec<JavaOverloadTarget>>,
    class_members: &std::collections::HashMap<String, JavaClassMemberNames>,
    current_class: Option<&str>,
    locals: &mut std::collections::HashSet<String>,
    local_types: &mut HashMap<String, String>,
) {
    for stmt in stmts {
        match &mut stmt.kind {
            StmtKind::VarDecl { declarations, .. } => {
                for decl in declarations {
                    if let Some(init) = &mut decl.init {
                        normalize_java_expr(
                            init,
                            fields,
                            methods,
                            static_methods,
                            static_overloads,
                            class_members,
                            current_class,
                            locals,
                            local_types,
                            false,
                        );
                        if let BindingPattern::Ident(name) = &decl.pattern {
                            if matches!(
                                decl.type_hint.as_deref(),
                                Some("Runnable" | "java.lang.Runnable")
                            ) {
                                JAVA_RUNNABLE_TARGETS.with(|targets| {
                                    targets.borrow_mut().insert(name.clone(), init.clone());
                                });
                                if java_thread_target_is_unsafe(init, locals) {
                                    JAVA_RUNNABLE_UNSAFE_TARGETS.with(|targets| {
                                        targets.borrow_mut().insert(name.clone());
                                    });
                                }
                            }
                            if let ExprKind::Call { callee, args, .. } = &init.kind {
                                if matches!(&callee.kind, ExprKind::Ident(c) if c == "__j_thread_new")
                                {
                                    if let Some(target) = args.first() {
                                        JAVA_THREAD_TARGETS.with(|targets| {
                                            targets
                                                .borrow_mut()
                                                .insert(name.clone(), target.value.clone());
                                        });
                                        if java_thread_target_is_unsafe(&target.value, locals) {
                                            JAVA_THREAD_UNSAFE_TARGETS.with(|targets| {
                                                targets.borrow_mut().insert(name.clone());
                                            });
                                        }
                                    }
                                }
                            }
                        }
                    }
                    collect_binding_names(&decl.pattern, locals);
                    let restored_type = match (&decl.pattern, decl.type_hint.as_deref()) {
                        (BindingPattern::Ident(name), None) => {
                            JAVA_LOCAL_TYPES.with(|types| types.borrow().get(name).cloned())
                        }
                        _ => None,
                    };
                    collect_binding_types(
                        &decl.pattern,
                        decl.type_hint.as_deref().or(restored_type.as_deref()),
                        local_types,
                    );
                }
            }
            StmtKind::Assign { targets, value } => {
                normalize_java_expr(
                    value,
                    fields,
                    methods,
                    static_methods,
                    static_overloads,
                    class_members,
                    current_class,
                    locals,
                    local_types,
                    false,
                );
                if let Some(target) = targets.first() {
                    if matches!(
                        target,
                        Expression {
                            kind: ExprKind::Ident(name),
                            ..
                        } if local_types.get(name).is_some_and(|ty| {
                            matches!(
                                java_type_simple_name(ty),
                                "byte" | "Byte" | "short" | "Short" | "int" | "Integer" | "long" | "Long"
                            )
                        })
                    ) && java_expr_is_char_numeric_source(value, local_types)
                    {
                        *value = java_cast_char_numeric_operand(value.clone(), local_types);
                    }
                }
                for target in &mut *targets {
                    normalize_java_expr(
                        target,
                        fields,
                        methods,
                        static_methods,
                        static_overloads,
                        class_members,
                        current_class,
                        locals,
                        local_types,
                        true,
                    );
                }
            }
            StmtKind::CompoundAssign { target, value, .. } => {
                normalize_java_expr(
                    value,
                    fields,
                    methods,
                    static_methods,
                    static_overloads,
                    class_members,
                    current_class,
                    locals,
                    local_types,
                    false,
                );
                normalize_java_expr(
                    target,
                    fields,
                    methods,
                    static_methods,
                    static_overloads,
                    class_members,
                    current_class,
                    locals,
                    local_types,
                    true,
                );
            }
            StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
                normalize_java_expr(
                    expr,
                    fields,
                    methods,
                    static_methods,
                    static_overloads,
                    class_members,
                    current_class,
                    locals,
                    local_types,
                    false,
                );
            }
            StmtKind::If {
                cond,
                then_body,
                elifs,
                else_body,
            } => {
                normalize_java_expr(
                    cond,
                    fields,
                    methods,
                    static_methods,
                    static_overloads,
                    class_members,
                    current_class,
                    locals,
                    local_types,
                    false,
                );
                normalize_java_stmts(
                    then_body,
                    fields,
                    methods,
                    static_methods,
                    static_overloads,
                    class_members,
                    current_class,
                    &mut locals.clone(),
                    &mut local_types.clone(),
                );
                for (elif_cond, elif_body) in elifs {
                    normalize_java_expr(
                        elif_cond,
                        fields,
                        methods,
                        static_methods,
                        static_overloads,
                        class_members,
                        current_class,
                        locals,
                        local_types,
                        false,
                    );
                    normalize_java_stmts(
                        elif_body,
                        fields,
                        methods,
                        static_methods,
                        static_overloads,
                        class_members,
                        current_class,
                        &mut locals.clone(),
                        &mut local_types.clone(),
                    );
                }
                if let Some(else_body) = else_body {
                    normalize_java_stmts(
                        else_body,
                        fields,
                        methods,
                        static_methods,
                        static_overloads,
                        class_members,
                        current_class,
                        &mut locals.clone(),
                        &mut local_types.clone(),
                    );
                }
            }
            StmtKind::For {
                init,
                cond,
                update,
                body,
            } => {
                let mut loop_locals = locals.clone();
                let mut loop_local_types = local_types.clone();
                if let Some(init_stmt) = init.as_mut() {
                    normalize_java_stmts(
                        std::slice::from_mut(init_stmt.as_mut()),
                        fields,
                        methods,
                        static_methods,
                        static_overloads,
                        class_members,
                        current_class,
                        &mut loop_locals,
                        &mut loop_local_types,
                    );
                }
                if let Some(cond) = cond {
                    normalize_java_expr(
                        cond,
                        fields,
                        methods,
                        static_methods,
                        static_overloads,
                        class_members,
                        current_class,
                        &loop_locals,
                        &loop_local_types,
                        false,
                    );
                }
                if let Some(update) = update {
                    normalize_java_expr(
                        update,
                        fields,
                        methods,
                        static_methods,
                        static_overloads,
                        class_members,
                        current_class,
                        &loop_locals,
                        &loop_local_types,
                        false,
                    );
                }
                normalize_java_stmts(
                    body,
                    fields,
                    methods,
                    static_methods,
                    static_overloads,
                    class_members,
                    current_class,
                    &mut loop_locals,
                    &mut loop_local_types,
                );
            }
            StmtKind::ForIn {
                var,
                key,
                iter,
                body,
                ..
            } => {
                normalize_java_expr(
                    iter,
                    fields,
                    methods,
                    static_methods,
                    static_overloads,
                    class_members,
                    current_class,
                    locals,
                    local_types,
                    false,
                );
                let mut loop_locals = locals.clone();
                let mut loop_local_types = local_types.clone();
                loop_locals.insert(var.clone());
                if let Some(type_hint) =
                    JAVA_LOCAL_TYPES.with(|types| types.borrow().get(var).cloned())
                {
                    loop_local_types.insert(var.clone(), type_hint);
                }
                if let Some(key) = key {
                    loop_locals.insert(key.clone());
                }
                normalize_java_stmts(
                    body,
                    fields,
                    methods,
                    static_methods,
                    static_overloads,
                    class_members,
                    current_class,
                    &mut loop_locals,
                    &mut loop_local_types,
                );
            }
            StmtKind::While { cond, body, .. } => {
                normalize_java_expr(
                    cond,
                    fields,
                    methods,
                    static_methods,
                    static_overloads,
                    class_members,
                    current_class,
                    locals,
                    local_types,
                    false,
                );
                normalize_java_stmts(
                    body,
                    fields,
                    methods,
                    static_methods,
                    static_overloads,
                    class_members,
                    current_class,
                    &mut locals.clone(),
                    &mut local_types.clone(),
                );
            }
            StmtKind::DoWhile { body, cond, .. } => {
                normalize_java_stmts(
                    body,
                    fields,
                    methods,
                    static_methods,
                    static_overloads,
                    class_members,
                    current_class,
                    &mut locals.clone(),
                    &mut local_types.clone(),
                );
                normalize_java_expr(
                    cond,
                    fields,
                    methods,
                    static_methods,
                    static_overloads,
                    class_members,
                    current_class,
                    locals,
                    local_types,
                    false,
                );
            }
            StmtKind::Switch {
                expr,
                cases,
                default,
            } => {
                normalize_java_expr(
                    expr,
                    fields,
                    methods,
                    static_methods,
                    static_overloads,
                    class_members,
                    current_class,
                    locals,
                    local_types,
                    false,
                );
                if java_expr_is_char_numeric_source(expr, local_types) {
                    *expr = java_cast_char_numeric_operand(expr.clone(), local_types);
                }
                for case in cases {
                    for condition in &mut case.conditions {
                        match condition {
                            CaseCondition::Value(value) => {
                                normalize_java_expr(
                                    value,
                                    fields,
                                    methods,
                                    static_methods,
                                    static_overloads,
                                    class_members,
                                    current_class,
                                    locals,
                                    local_types,
                                    false,
                                );
                                *value = java_char_numeric_cast_expr(value.clone());
                            }
                            CaseCondition::Range { from, to } => {
                                normalize_java_expr(
                                    from,
                                    fields,
                                    methods,
                                    static_methods,
                                    static_overloads,
                                    class_members,
                                    current_class,
                                    locals,
                                    local_types,
                                    false,
                                );
                                normalize_java_expr(
                                    to,
                                    fields,
                                    methods,
                                    static_methods,
                                    static_overloads,
                                    class_members,
                                    current_class,
                                    locals,
                                    local_types,
                                    false,
                                );
                            }
                            CaseCondition::Comparison { expr, .. } => {
                                normalize_java_expr(
                                    expr,
                                    fields,
                                    methods,
                                    static_methods,
                                    static_overloads,
                                    class_members,
                                    current_class,
                                    locals,
                                    local_types,
                                    false,
                                );
                            }
                        }
                    }
                    normalize_java_stmts(
                        &mut case.body,
                        fields,
                        methods,
                        static_methods,
                        static_overloads,
                        class_members,
                        current_class,
                        &mut locals.clone(),
                        &mut local_types.clone(),
                    );
                }
                if let Some(default) = default {
                    normalize_java_stmts(
                        default,
                        fields,
                        methods,
                        static_methods,
                        static_overloads,
                        class_members,
                        current_class,
                        &mut locals.clone(),
                        &mut local_types.clone(),
                    );
                }
            }
            StmtKind::Block(body) => {
                normalize_java_stmts(
                    body,
                    fields,
                    methods,
                    static_methods,
                    static_overloads,
                    class_members,
                    current_class,
                    &mut locals.clone(),
                    &mut local_types.clone(),
                );
            }
            StmtKind::Try {
                body,
                catches,
                else_body,
                finally,
            } => {
                normalize_java_stmts(
                    body,
                    fields,
                    methods,
                    static_methods,
                    static_overloads,
                    class_members,
                    current_class,
                    &mut locals.clone(),
                    &mut local_types.clone(),
                );
                for catch in catches {
                    let mut catch_locals = locals.clone();
                    if let Some(name) = &catch.var_name {
                        catch_locals.insert(name.clone());
                    }
                    normalize_java_stmts(
                        &mut catch.body,
                        fields,
                        methods,
                        static_methods,
                        static_overloads,
                        class_members,
                        current_class,
                        &mut catch_locals,
                        &mut local_types.clone(),
                    );
                }
                if let Some(else_body) = else_body {
                    normalize_java_stmts(
                        else_body,
                        fields,
                        methods,
                        static_methods,
                        static_overloads,
                        class_members,
                        current_class,
                        &mut locals.clone(),
                        &mut local_types.clone(),
                    );
                }
                if let Some(finally) = finally {
                    normalize_java_stmts(
                        finally,
                        fields,
                        methods,
                        static_methods,
                        static_overloads,
                        class_members,
                        current_class,
                        &mut locals.clone(),
                        &mut local_types.clone(),
                    );
                }
            }
            _ => {}
        }
    }
}

fn collect_binding_names(pattern: &BindingPattern, locals: &mut std::collections::HashSet<String>) {
    match pattern {
        BindingPattern::Ident(name) => {
            locals.insert(name.clone());
        }
        BindingPattern::Object(props) => {
            for prop in props {
                if let Some(value) = &prop.value {
                    collect_binding_names(value, locals);
                } else {
                    locals.insert(prop.key.clone());
                }
            }
        }
        BindingPattern::Array(elems) => {
            for elem in elems {
                match elem {
                    ArrayPatternElem::Pattern(pattern, _) => collect_binding_names(pattern, locals),
                    ArrayPatternElem::Rest(name) => {
                        locals.insert(name.clone());
                    }
                    ArrayPatternElem::Hole => {}
                }
            }
        }
    }
}

fn collect_binding_types(
    pattern: &BindingPattern,
    type_hint: Option<&str>,
    local_types: &mut HashMap<String, String>,
) {
    let Some(type_hint) = type_hint else {
        return;
    };
    match pattern {
        BindingPattern::Ident(name) => {
            local_types.insert(name.clone(), type_hint.to_string());
        }
        BindingPattern::Object(props) => {
            for prop in props {
                if let Some(value) = &prop.value {
                    collect_binding_types(value, Some(type_hint), local_types);
                }
            }
        }
        BindingPattern::Array(elems) => {
            for elem in elems {
                match elem {
                    ArrayPatternElem::Pattern(pattern, _) => {
                        collect_binding_types(pattern, Some(type_hint), local_types);
                    }
                    ArrayPatternElem::Rest(name) => {
                        local_types.insert(name.clone(), type_hint.to_string());
                    }
                    ArrayPatternElem::Hole => {}
                }
            }
        }
    }
}

fn normalize_java_expr(
    expr: &mut Expression,
    fields: &std::collections::HashSet<String>,
    methods: &std::collections::HashSet<String>,
    static_methods: &std::collections::HashSet<String>,
    static_overloads: &HashMap<String, Vec<JavaOverloadTarget>>,
    class_members: &std::collections::HashMap<String, JavaClassMemberNames>,
    current_class: Option<&str>,
    locals: &std::collections::HashSet<String>,
    local_types: &HashMap<String, String>,
    is_assignment_target: bool,
) {
    match &mut expr.kind {
        ExprKind::Ident(name) if fields.contains(name) && !locals.contains(name) => {
            let field = name.clone();
            expr.kind = ExprKind::Member {
                object: Box::new(Expression::new(ExprKind::This)),
                field,
                null_safe: false,
            };
        }
        ExprKind::Ident(name)
            if !locals.contains(name)
                && current_class.is_some()
                && JAVA_STATIC_FIELD_VARS.with(|vars| {
                    vars.borrow().contains(&format!(
                        "{}.{}",
                        current_class.unwrap_or_default(),
                        name
                    ))
                }) =>
        {
            let class_name = current_class.unwrap_or_default();
            let field = name.clone();
            expr.kind = ExprKind::Member {
                object: Box::new(Expression::ident(class_name)),
                field,
                null_safe: false,
            };
        }
        ExprKind::Call { callee, args, .. } => {
            for arg in args.iter_mut() {
                normalize_java_expr(
                    &mut arg.value,
                    fields,
                    methods,
                    static_methods,
                    static_overloads,
                    class_members,
                    current_class,
                    locals,
                    local_types,
                    false,
                );
            }
            if matches!(
                &callee.kind,
                ExprKind::Ident(name) if matches!(name.as_str(), "__j_print" | "__j_println" | "__java_print" | "__java_println")
            ) {
                if let Some(first) = args.get_mut(0) {
                    *first = java_print_arg(first.clone());
                }
            }
            if matches!(&callee.kind, ExprKind::Ident(name) if name == "__j_thread_start")
                && args.len() == 1
            {
                if let ExprKind::Ident(thread_name) = &args[0].value.kind {
                    if let Some(mut target) = JAVA_THREAD_TARGETS
                        .with(|targets| targets.borrow().get(thread_name).cloned())
                    {
                        target = java_resolve_runnable_target(target);
                        normalize_java_expr(
                            &mut target,
                            fields,
                            methods,
                            static_methods,
                            static_overloads,
                            class_members,
                            current_class,
                            locals,
                            local_types,
                            false,
                        );
                        java_rewrite_spawned_thread_sleep_expr(&mut target);
                        *expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident("__java_thread_start_with")),
                            args: vec![args[0].clone(), Argument::positional(target)],
                            optional: false,
                        });
                        return;
                    }
                }
            }
            if let ExprKind::Member { object, field, .. } = &mut callee.kind {
                if java_expr_dotted_name(object)
                    .as_deref()
                    .is_some_and(|name| java_type_simple_name(name) == "Executors")
                    && field == "newFixedThreadPool"
                {
                    *expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__j_exec_new")),
                        args: args.iter().cloned().collect(),
                        optional: false,
                    });
                    return;
                }
                if let Some(type_name) = java_expr_dotted_name(object) {
                    if let Some(class_names) = java_lookup_class_members(class_members, &type_name)
                    {
                        if let Some(target) =
                            select_java_overload_target(field, args, &class_names.static_overloads)
                        {
                            *field = target.mangled_name.clone();
                            if let Some(type_name) =
                                java_overload_return_cast_type(target.return_type.as_deref())
                            {
                                let inner = expr.clone();
                                *expr = Expression::new(ExprKind::Cast {
                                    type_name,
                                    expr: Box::new(inner),
                                });
                            }
                            return;
                        }
                    }
                }
                normalize_java_expr(
                    object,
                    fields,
                    methods,
                    static_methods,
                    static_overloads,
                    class_members,
                    current_class,
                    locals,
                    local_types,
                    false,
                );
                if let Some(receiver_type) = java_overload_receiver_type(object, local_types) {
                    if let Some(class_names) =
                        java_lookup_class_members(class_members, &receiver_type)
                    {
                        if let Some(target) = select_java_overload_target(
                            field,
                            args,
                            &class_names.instance_overloads,
                        ) {
                            *field = target.mangled_name.clone();
                            if let Some(type_name) =
                                java_overload_return_cast_type(target.return_type.as_deref())
                            {
                                let inner = expr.clone();
                                *expr = Expression::new(ExprKind::Cast {
                                    type_name,
                                    expr: Box::new(inner),
                                });
                            }
                            return;
                        }
                    }
                }
                if java_type_is_semaphore(java_receiver_type(object, local_types).as_deref()) {
                    if let Some(internal) = java_semaphore_method_name(field) {
                        let mut new_args = Vec::with_capacity(args.len() + 1);
                        new_args.push(Argument::positional((**object).clone()));
                        new_args.extend(args.iter().cloned());
                        *expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident(internal)),
                            args: new_args,
                            optional: false,
                        });
                        return;
                    }
                }
                if java_type_is_count_down_latch(java_receiver_type(object, local_types).as_deref())
                {
                    if let Some(internal) = java_count_down_latch_method_name(field, args.len()) {
                        let mut new_args = Vec::with_capacity(args.len() + 1);
                        new_args.push(Argument::positional((**object).clone()));
                        new_args.extend(args.iter().cloned());
                        *expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident(internal)),
                            args: new_args,
                            optional: false,
                        });
                        return;
                    }
                }
                if java_type_is_future_task(java_receiver_type(object, local_types).as_deref()) {
                    if let Some(internal) = java_future_task_method_name(field, args.len()) {
                        let mut new_args = Vec::with_capacity(args.len() + 1);
                        new_args.push(Argument::positional((**object).clone()));
                        new_args.extend(args.iter().cloned());
                        *expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident(internal)),
                            args: new_args,
                            optional: false,
                        });
                        return;
                    }
                }
                if java_type_is_executor_service(java_receiver_type(object, local_types).as_deref())
                {
                    if let Some(internal) = java_executor_method_name(field, args.len()) {
                        let mut new_args = Vec::with_capacity(args.len() + 1);
                        new_args.push(Argument::positional((**object).clone()));
                        new_args.extend(args.iter().cloned());
                        *expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident(internal)),
                            args: new_args,
                            optional: false,
                        });
                        return;
                    }
                }
                if java_type_is_queue_or_deque(java_receiver_type(object, local_types).as_deref()) {
                    let receiver_type = java_receiver_type(object, local_types);
                    if let Some(internal) =
                        java_queue_method_name(receiver_type.as_deref(), field, args.len())
                    {
                        let mut new_args = Vec::with_capacity(args.len() + 1);
                        new_args.push(Argument::positional((**object).clone()));
                        new_args.extend(args.iter().cloned());
                        *expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident(internal)),
                            args: new_args,
                            optional: false,
                        });
                        return;
                    }
                }
                if java_type_is_list_like(java_receiver_type(object, local_types).as_deref()) {
                    if let Some(internal) = java_list_method_name(field, args.len()) {
                        let mut new_args = Vec::with_capacity(args.len() + 1);
                        new_args.push(Argument::positional((**object).clone()));
                        new_args.extend(args.iter().cloned());
                        *expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident(internal)),
                            args: new_args,
                            optional: false,
                        });
                        return;
                    }
                }
                if java_type_is_spliterator(java_receiver_type(object, local_types).as_deref()) {
                    if let Some(internal) = java_spliterator_method_name(field, args.len()) {
                        let mut new_args = Vec::with_capacity(args.len() + 1);
                        new_args.push(Argument::positional((**object).clone()));
                        new_args.extend(args.iter().cloned());
                        *expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident(internal)),
                            args: new_args,
                            optional: false,
                        });
                        return;
                    }
                }
                if java_type_is_runtime(java_receiver_type(object, local_types).as_deref()) {
                    if let Some(internal) = java_runtime_method_name(field, args.len()) {
                        let mut new_args = Vec::with_capacity(args.len() + 1);
                        new_args.push(Argument::positional((**object).clone()));
                        new_args.extend(args.iter().cloned());
                        *expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident(internal)),
                            args: new_args,
                            optional: false,
                        });
                        return;
                    }
                }
                if java_type_is_process_builder(java_receiver_type(object, local_types).as_deref())
                {
                    if field == "command" && !args.is_empty() {
                        *expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident("__j_pb_command_set")),
                            args: vec![
                                Argument::positional((**object).clone()),
                                Argument::positional(java_args_to_array(args)),
                            ],
                            optional: false,
                        });
                        return;
                    }
                    if let Some(internal) = java_process_builder_method_name(field, args.len()) {
                        let mut new_args = Vec::with_capacity(args.len() + 1);
                        new_args.push(Argument::positional((**object).clone()));
                        new_args.extend(args.iter().cloned());
                        *expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident(internal)),
                            args: new_args,
                            optional: false,
                        });
                        return;
                    }
                }
                if java_type_is_process(java_receiver_type(object, local_types).as_deref()) {
                    if let Some(internal) = java_process_method_name(field, args.len()) {
                        let mut new_args = Vec::with_capacity(args.len() + 1);
                        new_args.push(Argument::positional((**object).clone()));
                        new_args.extend(args.iter().cloned());
                        *expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident(internal)),
                            args: new_args,
                            optional: false,
                        });
                        return;
                    }
                }
                if java_type_is_file(java_receiver_type(object, local_types).as_deref())
                    && field == "getPath"
                    && args.is_empty()
                {
                    *expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__j_file_get_path")),
                        args: vec![Argument::positional((**object).clone())],
                        optional: false,
                    });
                    return;
                }
                if java_type_is_redirect(java_receiver_type(object, local_types).as_deref())
                    && field == "type"
                    && args.is_empty()
                {
                    *expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__j_pb_redirect_type")),
                        args: vec![Argument::positional((**object).clone())],
                        optional: false,
                    });
                    return;
                }
            }
            if let ExprKind::Ident(name) = &callee.kind {
                if let Some(internal) = java_dotted_static_call_name(name) {
                    *callee = Box::new(Expression::ident(internal));
                    return;
                }
                if methods.contains(name) && !locals.contains(name) {
                    let method = name.clone();
                    callee.kind = ExprKind::Member {
                        object: Box::new(Expression::new(ExprKind::This)),
                        field: method,
                        null_safe: false,
                    };
                    return;
                }
                if static_methods.contains(name) && !locals.contains(name) {
                    if let Some(class_name) = current_class {
                        let target = select_java_overload_target(name, args, static_overloads);
                        let method = target
                            .map(|target| target.mangled_name.clone())
                            .unwrap_or_else(|| name.clone());
                        callee.kind = ExprKind::Member {
                            object: Box::new(Expression::ident(class_name)),
                            field: method,
                            null_safe: false,
                        };
                        if let Some(type_name) = target.and_then(|target| {
                            java_overload_return_cast_type(target.return_type.as_deref())
                        }) {
                            let inner = expr.clone();
                            *expr = Expression::new(ExprKind::Cast {
                                type_name,
                                expr: Box::new(inner),
                            });
                        }
                        return;
                    }
                }
            }
            normalize_java_expr(
                callee,
                fields,
                methods,
                static_methods,
                static_overloads,
                class_members,
                current_class,
                locals,
                local_types,
                false,
            );
        }
        ExprKind::Member { object, field, .. } => {
            if let Some(class_name) = java_expr_dotted_name(object) {
                if java_type_simple_name(&class_name) == "Spliterator" {
                    if let Some(value) = java_spliterator_constant(field) {
                        *expr = Expression::int(value);
                        return;
                    }
                }
                if class_name.ends_with("ProcessBuilder.Redirect") {
                    if matches!(field.as_str(), "PIPE" | "INHERIT" | "DISCARD") {
                        *expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident("__j_pb_redirect")),
                            args: vec![Argument::positional(Expression::string(field))],
                            optional: false,
                        });
                        return;
                    }
                }
            }
            if let Some(type_name) = java_qualified_static_type(object) {
                if type_name.contains('.') {
                    *object = Box::new(Expression::ident(&type_name));
                    return;
                }
            }
            normalize_java_expr(
                object,
                fields,
                methods,
                static_methods,
                static_overloads,
                class_members,
                current_class,
                locals,
                local_types,
                false,
            );
        }
        ExprKind::Index { object, index, .. } => {
            normalize_java_expr(
                object,
                fields,
                methods,
                static_methods,
                static_overloads,
                class_members,
                current_class,
                locals,
                local_types,
                false,
            );
            normalize_java_expr(
                index,
                fields,
                methods,
                static_methods,
                static_overloads,
                class_members,
                current_class,
                locals,
                local_types,
                false,
            );
        }
        ExprKind::Binary { op, left, right } => {
            normalize_java_expr(
                left,
                fields,
                methods,
                static_methods,
                static_overloads,
                class_members,
                current_class,
                locals,
                local_types,
                false,
            );
            normalize_java_expr(
                right,
                fields,
                methods,
                static_methods,
                static_overloads,
                class_members,
                current_class,
                locals,
                local_types,
                false,
            );
            if *op == BinOp::Add
                && (java_expr_is_string_value(left, local_types)
                    || java_expr_is_string_value(right, local_types))
            {
                let new_left = (**left).clone();
                let new_right = (**right).clone();
                *expr = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__java_string_concat")),
                    args: vec![
                        Argument::positional(new_left),
                        Argument::positional(new_right),
                    ],
                    optional: false,
                });
                return;
            }
            let has_char_numeric_operand = java_expr_is_char_numeric_source(left, local_types)
                || java_expr_is_char_numeric_source(right, local_types);
            if matches!(
                op,
                BinOp::Add
                    | BinOp::Sub
                    | BinOp::Mul
                    | BinOp::Div
                    | BinOp::Mod
                    | BinOp::BitAnd
                    | BinOp::BitOr
                    | BinOp::BitXor
                    | BinOp::Shl
                    | BinOp::Shr
                    | BinOp::UShr
            ) && (*op != BinOp::Add || has_char_numeric_operand)
            {
                *left = Box::new(java_cast_char_numeric_operand(
                    (**left).clone(),
                    local_types,
                ));
                *right = Box::new(java_cast_char_numeric_operand(
                    (**right).clone(),
                    local_types,
                ));
            }
        }
        ExprKind::Unary { op, expr: inner } => {
            normalize_java_expr(
                inner,
                fields,
                methods,
                static_methods,
                static_overloads,
                class_members,
                current_class,
                locals,
                local_types,
                is_assignment_target,
            );
            if matches!(
                op,
                UnaryOp::PreInc | UnaryOp::PostInc | UnaryOp::PreDec | UnaryOp::PostDec
            ) && matches!(&inner.kind, ExprKind::Ident(name) if local_types.get(name).is_some_and(|ty| java_type_simple_name(ty) == "char"))
            {
                let delta = if matches!(op, UnaryOp::PreDec | UnaryOp::PostDec) {
                    -1
                } else {
                    1
                };
                let target = (**inner).clone();
                let value = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__j_from_char_code")),
                    args: vec![Argument::positional(Expression::new(ExprKind::Binary {
                        op: BinOp::Add,
                        left: Box::new(Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident("__java_char_ord")),
                            args: vec![Argument::positional(target.clone())],
                            optional: false,
                        })),
                        right: Box::new(Expression::int(delta)),
                    }))],
                    optional: false,
                });
                *expr = Expression::new(ExprKind::Assign {
                    target: Box::new(target),
                    value: Box::new(value),
                });
                return;
            }
            rewrite_java_this_field_update(expr);
        }
        ExprKind::Assign { target, value } => {
            normalize_java_expr(
                value,
                fields,
                methods,
                static_methods,
                static_overloads,
                class_members,
                current_class,
                locals,
                local_types,
                false,
            );
            if matches!(
                &target.kind,
                ExprKind::Ident(name)
                    if local_types.get(name).is_some_and(|ty| {
                        matches!(
                            java_type_simple_name(ty),
                            "byte" | "Byte" | "short" | "Short" | "int" | "Integer" | "long" | "Long"
                        )
                    })
            ) && java_expr_is_char_numeric_source(value, local_types)
            {
                **value = java_cast_char_numeric_operand((**value).clone(), local_types);
            }
            normalize_java_expr(
                target,
                fields,
                methods,
                static_methods,
                static_overloads,
                class_members,
                current_class,
                locals,
                local_types,
                true,
            );
            rewrite_java_this_field_assign(expr);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            normalize_java_expr(
                cond,
                fields,
                methods,
                static_methods,
                static_overloads,
                class_members,
                current_class,
                locals,
                local_types,
                false,
            );
            normalize_java_expr(
                then,
                fields,
                methods,
                static_methods,
                static_overloads,
                class_members,
                current_class,
                locals,
                local_types,
                false,
            );
            normalize_java_expr(
                else_,
                fields,
                methods,
                static_methods,
                static_overloads,
                class_members,
                current_class,
                locals,
                local_types,
                false,
            );
        }
        ExprKind::Array(elems) => {
            for elem in elems {
                normalize_java_expr(
                    &mut elem.value,
                    fields,
                    methods,
                    static_methods,
                    static_overloads,
                    class_members,
                    current_class,
                    locals,
                    local_types,
                    false,
                );
            }
        }
        ExprKind::Lambda { params, body, .. } => {
            let mut lambda_locals = locals.clone();
            let mut lambda_types = local_types.clone();
            for param in params {
                lambda_locals.insert(param.name.clone());
                if let Some(type_hint) = &param.type_hint {
                    lambda_types.insert(param.name.clone(), type_hint.clone());
                }
            }
            match body {
                LambdaBody::Expr(inner) => normalize_java_expr(
                    inner,
                    fields,
                    methods,
                    static_methods,
                    static_overloads,
                    class_members,
                    current_class,
                    &lambda_locals,
                    &lambda_types,
                    false,
                ),
                LambdaBody::Block(stmts) => normalize_java_stmts(
                    stmts,
                    fields,
                    methods,
                    static_methods,
                    static_overloads,
                    class_members,
                    current_class,
                    &mut lambda_locals,
                    &mut lambda_types,
                ),
            }
        }
        ExprKind::StaticAccess { class, member } => {
            normalize_java_expr(
                class,
                fields,
                methods,
                static_methods,
                static_overloads,
                class_members,
                current_class,
                locals,
                local_types,
                false,
            );
            normalize_java_expr(
                member,
                fields,
                methods,
                static_methods,
                static_overloads,
                class_members,
                current_class,
                locals,
                local_types,
                false,
            );
            if let (Some(class_name), ExprKind::Ident(member_name)) =
                (java_expr_dotted_name(class), &member.kind)
            {
                let member_name = member_name.clone();
                if java_type_simple_name(&class_name) == "Spliterator" {
                    if let Some(value) = java_spliterator_constant(&member_name) {
                        *expr = Expression::int(value);
                    }
                }
                if class_name.ends_with("ProcessBuilder.Redirect")
                    && matches!(member_name.as_str(), "PIPE" | "INHERIT" | "DISCARD")
                {
                    *expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__j_pb_redirect")),
                        args: vec![Argument::positional(Expression::string(&member_name))],
                        optional: false,
                    });
                }
            }
        }
        _ => {}
    }
}

fn java_this_member(expr: &Expression) -> Option<(Expression, String)> {
    let _ = expr;
    None
}

fn rewrite_java_this_field_assign(expr: &mut Expression) {
    let ExprKind::Assign { target, value } = &expr.kind else {
        return;
    };
    let Some((object, field)) = java_this_member(target) else {
        return;
    };
    *expr = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("__java_field_set")),
        args: vec![
            Argument::positional(object),
            Argument::positional(Expression::string(&field)),
            Argument::positional((**value).clone()),
        ],
        optional: false,
    });
}

fn rewrite_java_this_field_update(expr: &mut Expression) {
    let ExprKind::Unary { op, expr: inner } = &expr.kind else {
        return;
    };
    let delta = match op {
        UnaryOp::PreInc | UnaryOp::PostInc => 1,
        UnaryOp::PreDec | UnaryOp::PostDec => -1,
        _ => return,
    };
    let Some((object, field)) = java_this_member(inner) else {
        return;
    };
    *expr = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("__java_field_inc")),
        args: vec![
            Argument::positional(object),
            Argument::positional(Expression::string(&field)),
            Argument::positional(Expression::int(delta)),
        ],
        optional: false,
    });
}

// ════════════════════════════════════════════════════════════════════════════
// Helpers
// ════════════════════════════════════════════════════════════════════════════

/// Extract a simple type name from a `type_ref` or `ref_type` node.
fn extract_ref_name(pair: &Pair<Rule>) -> String {
    match pair.as_rule() {
        Rule::type_ref => {
            let dims = pair
                .clone()
                .into_inner()
                .filter(|p| p.as_rule() == Rule::dim_suffix)
                .count();
            for p in pair.clone().into_inner() {
                match p.as_rule() {
                    Rule::primitive_type => return format!("{}{}", p.as_str(), "[]".repeat(dims)),
                    Rule::ref_type => {
                        return format!("{}{}", extract_ref_name(&p), "[]".repeat(dims));
                    }
                    _ => {}
                }
            }
            let raw = pair.as_str().trim().trim_end_matches("[]").trim();
            let base = common_generics::parse_type_ref_hint(raw)
                .map(|ty| common_generics::display_type_ref(&ty))
                .unwrap_or_else(|| common_generics::generic_base_name(raw).to_string());
            format!("{}{}", base, "[]".repeat(dims))
        }
        Rule::ref_type => {
            let raw = pair.as_str().trim();
            common_generics::parse_type_ref_hint(raw)
                .map(|ty| common_generics::display_type_ref(&ty))
                .unwrap_or_else(|| common_generics::generic_base_name(raw).to_string())
        }
        _ => {
            let raw = pair.as_str().trim();
            common_generics::parse_type_ref_hint(raw)
                .map(|ty| common_generics::display_type_ref(&ty))
                .unwrap_or_else(|| common_generics::generic_base_name(raw).to_string())
        }
    }
}

fn str_to_binop(s: &str) -> BinOp {
    match s {
        "+" => BinOp::Add,
        "-" => BinOp::Sub,
        "*" => BinOp::Mul,
        "/" => BinOp::Div,
        "%" => BinOp::Mod,
        "==" => BinOp::Eq,
        "!=" => BinOp::NotEq,
        "<" => BinOp::Lt,
        "<=" => BinOp::LtEq,
        ">" => BinOp::Gt,
        ">=" => BinOp::GtEq,
        "&&" => BinOp::And,
        "||" => BinOp::Or,
        "&" => BinOp::BitAnd,
        "|" => BinOp::BitOr,
        "^" => BinOp::BitXor,
        "<<" => BinOp::Shl,
        ">>" => BinOp::Shr,
        ">>>" => BinOp::UShr,
        _ => BinOp::Add,
    }
}

fn compound_op_to_binop(s: &str) -> BinOp {
    match s.trim_end_matches('=') {
        "+" => BinOp::Add,
        "-" => BinOp::Sub,
        "*" => BinOp::Mul,
        "/" => BinOp::Div,
        "%" => BinOp::Mod,
        "&" => BinOp::BitAnd,
        "|" => BinOp::BitOr,
        "^" => BinOp::BitXor,
        "<<" => BinOp::Shl,
        ">>" => BinOp::Shr,
        ">>>" => BinOp::UShr,
        _ => BinOp::Add,
    }
}

/// JLS §3.10.6 text-block content: strip the opening line, remove incidental
/// leading whitespace (minimum over non-blank lines plus the closing
/// delimiter's own line), strip incidental trailing whitespace per line,
/// keep the final newline when the closing delimiter sits on its own line,
/// then process escape sequences.
fn java_text_block_content(raw: &str) -> String {
    // The opening line holds only optional whitespace before its terminator.
    let raw = match raw.find('\n') {
        Some(pos) => &raw[pos + 1..],
        None => raw,
    };
    let lines: Vec<&str> = raw.split('\n').collect();
    let last_index = lines.len().saturating_sub(1);
    let closing_on_own_line = lines
        .last()
        .is_some_and(|line| line.trim_start_matches([' ', '\t']).is_empty());

    let mut min_indent = usize::MAX;
    for (i, line) in lines.iter().enumerate() {
        let content = line.trim_start_matches([' ', '\t']);
        if i == last_index && closing_on_own_line {
            // Whitespace before the closing delimiter counts as indentation.
            min_indent = min_indent.min(line.len());
        } else if !content.is_empty() {
            min_indent = min_indent.min(line.len() - content.len());
        }
    }
    if min_indent == usize::MAX {
        min_indent = 0;
    }

    let mut out = String::new();
    for (i, line) in lines.iter().enumerate() {
        if i == last_index && closing_on_own_line {
            break;
        }
        let stripped = if line.len() >= min_indent {
            &line[min_indent..]
        } else {
            line.trim_start_matches([' ', '\t'])
        };
        out.push_str(stripped.trim_end_matches([' ', '\t']));
        if i < last_index {
            out.push('\n');
        }
    }
    unescape_java_string(&out)
}

fn unescape_java_string(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    let mut chars = s.chars();
    while let Some(c) = chars.next() {
        if c == '\\' {
            match chars.next() {
                Some('u') => {
                    let mut hex = String::with_capacity(4);
                    while matches!(chars.clone().next(), Some('u')) {
                        chars.next();
                    }
                    for _ in 0..4 {
                        if let Some(h) = chars.next() {
                            hex.push(h);
                        }
                    }
                    if let Ok(code) = u32::from_str_radix(&hex, 16) {
                        if (0xD800..0xDC00).contains(&code) {
                            // High surrogate: combine with the following
                            // \uXXXX low surrogate (JLS §3.10.6 pairs) —
                            // char::from_u32 rejects lone surrogates and
                            // silently DROPPED the whole pair before.
                            let mut peek = chars.clone();
                            if peek.next() == Some('\\') && peek.next() == Some('u') {
                                let mut lo_hex = String::with_capacity(4);
                                for _ in 0..4 {
                                    if let Some(h) = peek.next() {
                                        lo_hex.push(h);
                                    }
                                }
                                if let Ok(lo) = u32::from_str_radix(&lo_hex, 16) {
                                    if (0xDC00..0xE000).contains(&lo) {
                                        let cp = 0x10000 + ((code - 0xD800) << 10) + (lo - 0xDC00);
                                        if let Some(ch) = char::from_u32(cp) {
                                            out.push(ch);
                                            chars = peek;
                                        }
                                    }
                                }
                            }
                        } else if let Some(ch) = char::from_u32(code) {
                            out.push(ch);
                        }
                    }
                }
                Some('n') => out.push('\n'),
                Some('t') => out.push('\t'),
                Some('r') => out.push('\r'),
                Some('"') => out.push('"'),
                Some('\'') => out.push('\''),
                Some('\\') => out.push('\\'),
                Some('0') => out.push('\0'),
                Some(c) => {
                    out.push('\\');
                    out.push(c);
                }
                None => {}
            }
        } else {
            out.push(c);
        }
    }
    out
}

fn java_unicode_escape_code_unit(s: &str) -> Option<u32> {
    let rest = s.strip_prefix("\\u")?;
    if rest.len() < 4 {
        return None;
    }
    u32::from_str_radix(&rest[..4], 16).ok()
}

fn find_main_class(body: &[Statement]) -> Option<String> {
    for stmt in body {
        if let StmtKind::ClassDecl { name, members, .. } = &stmt.kind {
            for m in members {
                if let ClassMember::Method(func) = m {
                    if let StmtKind::FunctionDecl {
                        name: fname,
                        modifiers,
                        ..
                    } = &func.kind
                    {
                        if fname == "main" && modifiers.is_static {
                            return Some(name.clone());
                        }
                    }
                }
            }
        }
    }
    None
}
