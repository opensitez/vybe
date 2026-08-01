use crate::emitter::tostring::SET_MARKER;
use pest::Parser;
use pest::iterators::Pair;
use std::collections::{HashMap, HashSet};
use std::sync::atomic::{AtomicUsize, Ordering};
use vybe_ast::*;

use super::{KotlinParser, Rule};

static NEXT_TMP_ID: AtomicUsize = AtomicUsize::new(1);

fn gen_tmp_name() -> String {
    format!("__kt_tmp_{}", NEXT_TMP_ID.fetch_add(1, Ordering::Relaxed))
}

thread_local! {
    /// Every method name some class in this source declares.
    static USER_MEMBER_NAMES: std::cell::RefCell<std::collections::HashSet<(String, usize)>> =
        std::cell::RefCell::new(std::collections::HashSet::new());
    static USER_METHOD_OVERLOADS: std::cell::RefCell<std::collections::HashMap<String, std::collections::HashMap<usize, Vec<Vec<String>>>>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
    /// Every property name some class in this source declares — the same
    /// problem one level down: `.values`, `.size`, `.keys`, `.first` are
    /// rewritten to dict/collection primitives on SPELLING, so a data class
    /// with `val values: MutableList<Int>` had `a.values` return the OBJECT's
    /// members instead of the list.
    static USER_PROPERTY_NAMES: std::cell::RefCell<std::collections::HashSet<String>> =
        std::cell::RefCell::new(std::collections::HashSet::new());
    /// Source-declared types. These shadow stdlib constructor helpers such as
    /// `Pair(...)` / `Triple(...)`; a user data class named `Pair` must still
    /// normalize to `New(Pair, ...)`, not to the tuple literal helper.
    static USER_CLASS_NAMES: std::cell::RefCell<std::collections::HashSet<String>> =
        std::cell::RefCell::new(std::collections::HashSet::new());
    /// Names bound by each enclosing function being walked — its parameters and
    /// everything its body declares.
    static ENCLOSING_LOCALS: std::cell::RefCell<Vec<std::collections::HashSet<String>>> =
        const { std::cell::RefCell::new(Vec::new()) };
    /// Local class -> the enclosing locals it reads. Kotlin gives a local class
    /// synthetic storage for the values it captures; declaring them as leading
    /// constructor parameters is that lowering, and it is the frontend's job.
    static LOCAL_CLASS_CAPTURES: std::cell::RefCell<std::collections::HashMap<String, Vec<String>>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
    /// The superclass of each class whose body is being walked. `super<X>.f()`
    /// means the PARENT's `f` when `X` is that parent, and an interface
    /// DEFAULT when it is not — two different lowerings.
    static CURRENT_CLASS_PARENT: std::cell::RefCell<Vec<Option<String>>> =
        const { std::cell::RefCell::new(Vec::new()) };
    /// `(interface, member)` pairs reached as `super<I>.m(...)` in the class
    /// being walked. Each becomes an additive `AugmentAdjustment` alias on that
    /// interface's augmentation, which is how the shared fold already exposes a
    /// contributed member under a second name (PHP `use A { m as alias; }`).
    static SUPER_QUALIFIED_USES: std::cell::RefCell<Vec<Vec<(String, String)>>> =
        const { std::cell::RefCell::new(Vec::new()) };
    /// Top-level `fun Receiver.name(...)` declarations. An extension is a plain
    /// function whose first parameter is the receiver, so `x.name(a)` has to be
    /// rewritten to `name(x, a)` — the same lowering VB's walker applies to
    /// `<Extension()>` methods.
    static EXTENSION_FUNCTIONS: std::cell::RefCell<std::collections::HashSet<String>> =
        std::cell::RefCell::new(std::collections::HashSet::new());
    static EXTENSION_PROPERTIES: std::cell::RefCell<std::collections::HashSet<String>> =
        std::cell::RefCell::new(std::collections::HashSet::new());
    /// Class name -> the members it declares. An extension body reads the
    /// receiver's members unqualified (`val P.twice get() = n * 2`), so those
    /// names have to resolve to `this.<name>`.
    static CLASS_MEMBERS: std::cell::RefCell<std::collections::HashMap<String, std::collections::HashSet<String>>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
    /// The receiver's members, while an extension declaration's body is walked.
    static EXT_RECEIVER_MEMBERS: std::cell::RefCell<Vec<std::collections::HashSet<String>>> =
        const { std::cell::RefCell::new(Vec::new()) };
    static CURRENT_CLASS_STACK: std::cell::RefCell<Vec<String>> =
        const { std::cell::RefCell::new(Vec::new()) };
    static INNER_CLASS_QUALIFIED: std::cell::RefCell<std::collections::HashMap<String, String>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
    static INNER_OUTER_MEMBERS: std::cell::RefCell<Vec<Vec<(String, std::collections::HashSet<String>)>>> =
        const { std::cell::RefCell::new(Vec::new()) };
}

fn outer_this(depth: usize) -> Expression {
    let mut expr = Expression::new(ExprKind::This);
    for _ in 0..depth {
        expr = Expression::new(ExprKind::Member {
            object: Box::new(expr),
            field: "__kt_outer".to_string(),
            null_safe: false,
        });
    }
    expr
}

fn current_local_has(name: &str) -> bool {
    ENCLOSING_LOCALS.with(|stack| stack.borrow().last().is_some_and(|s| s.contains(name)))
}

fn inner_outer_read(name: &str) -> Option<Expression> {
    if current_local_has(name) {
        return None;
    }
    INNER_OUTER_MEMBERS.with(|stack| {
        stack.borrow().last().and_then(|outers| {
            outers
                .iter()
                .enumerate()
                .find(|(_, (_, members))| members.contains(name))
                .map(|(idx, _)| {
                    Expression::new(ExprKind::Member {
                        object: Box::new(outer_this(idx + 1)),
                        field: name.to_string(),
                        null_safe: false,
                    })
                })
        })
    })
}

fn qualified_inner_class(name: &str) -> Option<String> {
    INNER_CLASS_QUALIFIED.with(|m| m.borrow().get(name).cloned())
}

fn qualified_type_expr(path: &str) -> Expression {
    let mut parts = path.split('.');
    let Some(first) = parts.next() else {
        return Expression::ident(path);
    };
    let mut expr = Expression::ident(first);
    for part in parts {
        expr = Expression::new(ExprKind::Member {
            object: Box::new(expr),
            field: part.to_string(),
            null_safe: false,
        });
    }
    expr
}

/// `name` resolved inside an extension body: a member of the receiver becomes
/// `this.<name>`, everything else stays itself.
fn extension_receiver_read(name: &str) -> Option<Expression> {
    EXT_RECEIVER_MEMBERS.with(|stack| {
        stack.borrow().last().and_then(|members| {
            members.contains(name).then(|| {
                Expression::new(ExprKind::Member {
                    object: Box::new(Expression::new(ExprKind::This)),
                    field: name.to_string(),
                    null_safe: false,
                })
            })
        })
    })
}

/// Push the receiver's member set for the duration of an extension body.
fn push_ext_receiver(receiver: &str) {
    let members = CLASS_MEMBERS
        .with(|m| m.borrow().get(receiver).cloned())
        .unwrap_or_default();
    EXT_RECEIVER_MEMBERS.with(|stack| stack.borrow_mut().push(members));
}

fn pop_ext_receiver() {
    EXT_RECEIVER_MEMBERS.with(|stack| {
        stack.borrow_mut().pop();
    });
}

/// Top-level `val Receiver.name get() = …`. Read as a member (`x.name`), so
/// the READ — not a call — becomes `name(x)`.
fn is_extension_property(name: &str) -> bool {
    EXTENSION_PROPERTIES.with(|set| set.borrow().contains(name))
}

/// Whether `name` is declared as an extension function in this source.
fn is_extension_function(name: &str) -> bool {
    EXTENSION_FUNCTIONS.with(|set| set.borrow().contains(name))
}

/// The name a supertype's member is additionally bound under so `super<I>.m()`
/// can reach it after the class's own `m` has shadowed the default.
fn super_alias(from: &str, member: &str) -> String {
    format!("__super_{}_{}", from.replace('.', "_"), member)
}

/// Every identifier `pair`'s subtree binds: parameters, `val`/`var` names.
fn bound_names(pair: &Pair<Rule>) -> std::collections::HashSet<String> {
    fn walk(pair: Pair<Rule>, out: &mut std::collections::HashSet<String>) {
        match pair.as_rule() {
            Rule::parameter | Rule::class_parameter | Rule::var_decl | Rule::destructuring_decl => {
                for sub in pair.clone().into_inner() {
                    if sub.as_rule() == Rule::identifier {
                        out.insert(sub.as_str().to_string());
                    }
                    if sub.as_rule() == Rule::destructuring_target {
                        for id in sub.into_inner() {
                            if id.as_rule() == Rule::identifier {
                                out.insert(id.as_str().to_string());
                            }
                        }
                    }
                }
            }
            _ => {}
        }
        for sub in pair.into_inner() {
            walk(sub, out);
        }
    }
    let mut out = std::collections::HashSet::new();
    walk(pair.clone(), &mut out);
    out
}

/// Every identifier `pair`'s subtree READS.
fn read_names(pair: &Pair<Rule>) -> std::collections::HashSet<String> {
    fn walk(pair: Pair<Rule>, out: &mut std::collections::HashSet<String>) {
        if pair.as_rule() == Rule::identifier {
            out.insert(pair.as_str().to_string());
        }
        for sub in pair.into_inner() {
            walk(sub, out);
        }
    }
    let mut out = std::collections::HashSet::new();
    walk(pair.clone(), &mut out);
    out
}

/// Record the methods the source's own classes declare.
///
/// The collection rewrites below turn `x.add(v)` into `__coll_push(x, v)` on
/// SPELLING alone — they cannot see the receiver's type. That stole every user
/// method sharing one of ~20 collection names: `class Calc { fun add(x: Int) =
/// base + x }` had `c.add(3)` compiled as an array push and answered `1`.
///
/// A member declared by a class in this program always wins, which is Kotlin's
/// own rule (a member beats an extension). Anything else keeps the rewrite.
fn collect_user_member_names(root: &Pair<Rule>) {
    fn walk(
        pair: Pair<Rule>,
        out: &mut std::collections::HashSet<(String, usize)>,
        overloads: &mut std::collections::HashMap<
            String,
            std::collections::HashMap<usize, Vec<Vec<String>>>,
        >,
        in_class: bool,
        owner: Option<&str>,
    ) {
        let rule = pair.as_rule();
        // The class this subtree belongs to, so its members can be recorded
        // against it by name.
        let owner_here = if matches!(rule, Rule::class_decl | Rule::interface_decl) {
            pair.clone()
                .into_inner()
                .find(|p| p.as_rule() == Rule::identifier)
                .map(|p| p.as_str().to_string())
        } else {
            None
        };
        let qualified_owner_here = owner_here.as_ref().map(|class_name| {
            owner
                .map(|enclosing| format!("{enclosing}.{class_name}"))
                .unwrap_or_else(|| class_name.clone())
        });
        if let Some(class_name) = &owner_here {
            USER_CLASS_NAMES.with(|set| {
                set.borrow_mut().insert(class_name.clone());
            });
            let is_inner = pair
                .clone()
                .into_inner()
                .any(|p| p.as_rule() == Rule::modifier && p.as_str().trim() == "inner");
            if is_inner {
                if let Some(enclosing) = owner {
                    INNER_CLASS_QUALIFIED.with(|m| {
                        m.borrow_mut()
                            .insert(class_name.clone(), format!("{enclosing}.{class_name}"));
                    });
                }
            }
        }
        let owner_ref: Option<&str> = qualified_owner_here.as_deref().or(owner);
        let record_member = |name: &str| {
            if let Some(owner) = owner_ref {
                CLASS_MEMBERS.with(|m| {
                    m.borrow_mut()
                        .entry(owner.to_string())
                        .or_default()
                        .insert(name.to_string())
                });
            }
        };
        if rule == Rule::function_decl && in_class {
            let inners: Vec<_> = pair.clone().into_inner().collect();
            let name = inners
                .iter()
                .find(|p| p.as_rule() == Rule::identifier)
                .map(|p| p.as_str().to_string());
            // ARITY too, not the name alone. `interface Summer { fun sum(v:
            // IntArray): Int }` would otherwise disable the zero-argument
            // `values.sum()` inside its own implementation.
            let arity = inners
                .iter()
                .find(|p| p.as_rule() == Rule::parameter_list)
                .map(|p| {
                    p.clone()
                        .into_inner()
                        .filter(|q| q.as_rule() == Rule::parameter)
                        .count()
                })
                .unwrap_or(0);
            if let Some(name) = name {
                record_member(&name);
                let param_types = inners
                    .iter()
                    .find(|p| p.as_rule() == Rule::parameter_list)
                    .map(|p| {
                        p.clone()
                            .into_inner()
                            .filter(|q| q.as_rule() == Rule::parameter)
                            .map(|param| {
                                param
                                    .into_inner()
                                    .find(|q| q.as_rule() == Rule::type_ref)
                                    .map(|q| type_hint_text(q.as_str()))
                                    .unwrap_or_default()
                            })
                            .collect::<Vec<_>>()
                    })
                    .unwrap_or_default();
                overloads
                    .entry(name.clone())
                    .or_default()
                    .entry(arity)
                    .or_default()
                    .push(param_types);
                out.insert((name, arity));
            }
        }
        if rule == Rule::function_decl {
            let inners: Vec<_> = pair.clone().into_inner().collect();
            if inners.iter().any(|p| p.as_rule() == Rule::receiver_prefix) {
                if let Some(id) = inners.iter().find(|p| p.as_rule() == Rule::identifier) {
                    EXTENSION_FUNCTIONS
                        .with(|set| set.borrow_mut().insert(id.as_str().to_string()));
                }
            }
        }
        if rule == Rule::var_decl {
            let inners: Vec<_> = pair.clone().into_inner().collect();
            if inners.iter().any(|p| p.as_rule() == Rule::receiver_prefix) {
                if let Some(id) = inners.iter().find(|p| p.as_rule() == Rule::identifier) {
                    EXTENSION_PROPERTIES
                        .with(|set| set.borrow_mut().insert(id.as_str().to_string()));
                }
            }
        }
        if in_class && matches!(rule, Rule::var_decl | Rule::class_parameter) {
            // A `class_parameter` is only a property when it says `val`/`var`.
            let inners: Vec<_> = pair.clone().into_inner().collect();
            let is_property = rule == Rule::var_decl
                || inners
                    .iter()
                    .any(|p| matches!(p.as_rule(), Rule::val_kw | Rule::var_kw));
            if is_property {
                if let Some(id) = inners.iter().find(|p| p.as_rule() == Rule::identifier) {
                    record_member(id.as_str());
                    USER_PROPERTY_NAMES
                        .with(|set| set.borrow_mut().insert(id.as_str().to_string()));
                }
            }
        }
        // A primary constructor's `val`/`var` parameters are PROPERTIES of the
        // class even though they sit outside its body — `data class
        // Counter(val values: MutableList<Int>)` declares `values`.
        let nested = in_class || matches!(rule, Rule::class_body | Rule::primary_constructor);
        for sub in pair.into_inner() {
            walk(sub, out, overloads, nested, owner_ref);
        }
    }
    USER_PROPERTY_NAMES.with(|set| set.borrow_mut().clear());
    USER_CLASS_NAMES.with(|set| set.borrow_mut().clear());
    EXTENSION_FUNCTIONS.with(|set| set.borrow_mut().clear());
    EXTENSION_PROPERTIES.with(|set| set.borrow_mut().clear());
    CLASS_MEMBERS.with(|m| m.borrow_mut().clear());
    USER_METHOD_OVERLOADS.with(|map| map.borrow_mut().clear());
    INNER_CLASS_QUALIFIED.with(|m| m.borrow_mut().clear());
    USER_MEMBER_NAMES.with(|set| {
        let mut set = set.borrow_mut();
        set.clear();
        let mut overloads = std::collections::HashMap::new();
        walk(root.clone(), &mut set, &mut overloads, false, None);
        overloads.retain(|_, by_arity| by_arity.values().map(Vec::len).sum::<usize>() > 1);
        USER_METHOD_OVERLOADS.with(|map| {
            *map.borrow_mut() = overloads;
        });
    });
}

/// Whether a class in this source declares a PROPERTY with this name.
fn is_user_property_name(name: &str) -> bool {
    USER_PROPERTY_NAMES.with(|set| set.borrow().contains(name))
}

fn is_user_class_name(name: &str) -> bool {
    USER_CLASS_NAMES.with(|set| set.borrow().contains(name))
}

/// Whether a class in this source declares a method with this name and arity.
fn is_user_member_name(name: &str, arity: usize) -> bool {
    USER_MEMBER_NAMES.with(|set| set.borrow().contains(&(name.to_string(), arity)))
}

fn overloaded_storage_name(name: &str, arity: usize) -> Option<String> {
    USER_METHOD_OVERLOADS.with(|map| {
        let map = map.borrow();
        let signatures = map.get(name)?.get(&arity)?;
        if signatures.len() != 1 {
            return None;
        }
        let param_types = &signatures[0];
        if param_types.is_empty() {
            Some(format!("{name}$sig0"))
        } else if param_types.iter().all(|ty| !ty.is_empty()) {
            Some(format!("{name}$sig{}", param_types.join("$")))
        } else {
            None
        }
    })
}

pub fn parse(source: &str) -> Result<Module, String> {
    let mut pairs = KotlinParser::parse(Rule::program, source)
        .map_err(|e| format!("Kotlin parse error: {}", e))?;

    let root = pairs
        .next()
        .ok_or_else(|| "Empty parse result".to_string())?;
    collect_user_member_names(&root);
    let mut body = Vec::new();
    let imports = collect_imports(&root);
    let mut package_name: Option<String> = None;

    for pair in root.into_inner() {
        match pair.as_rule() {
            Rule::package_decl => {
                package_name = pair
                    .into_inner()
                    .find(|p| p.as_rule() == Rule::dotted_name)
                    .map(|p| p.as_str().to_string());
            }
            Rule::import_decl => {
                // Already collected recursively above. Imports are declarations
                // for resolver metadata, not executable statements.
            }
            Rule::top_level_decl => {
                let mut label_name = None;
                for inner in pair.into_inner() {
                    match inner.as_rule() {
                        Rule::label_decl => {
                            label_name = Some(inner.as_str().trim_end_matches('@').to_string())
                        }
                        Rule::typealias_decl => {
                            if let Some(stmt) = walk_typealias(inner) {
                                body.push(stmt);
                            }
                        }
                        _ => {
                            if let Some(stmt) = walk_statement(inner) {
                                if let Some(lbl) = label_name.take() {
                                    body.push(Statement::new(StmtKind::Labeled {
                                        label: lbl,
                                        body: Box::new(stmt),
                                    }));
                                } else {
                                    body.push(stmt);
                                }
                            }
                        }
                    }
                }
            }
            Rule::EOI => {}
            _ => {}
        }
    }

    let aliases = kotlin_import_aliases(&imports);
    rewrite_import_aliases_in_stmts(&mut body, &aliases);
    normalize_kotlin_operator_calls(&mut body);

    if let Some(name) = package_name.filter(|name| !name.is_empty()) {
        let has_main = body.iter().any(|stmt| {
            matches!(
                &stmt.kind,
                StmtKind::FunctionDecl { name, params, .. } if name == "main" && params.is_empty()
            )
        });
        body = vec![Statement::new(StmtKind::NamespaceDecl { name, body })];
        if has_main {
            body.push(Statement::new(StmtKind::Expr(Expression::new(
                ExprKind::Call {
                    callee: Box::new(Expression::ident("main")),
                    args: Vec::new(),
                    optional: false,
                },
            ))));
        }
    }

    Ok(Module {
        name: "main".to_string(),
        language: Lang::Kotlin,
        body,
        imports,
    })
}

fn collect_imports(root: &Pair<Rule>) -> Vec<Import> {
    fn walk(pair: Pair<Rule>, imports: &mut Vec<Import>, seen: &mut HashSet<String>) {
        if pair.as_rule() == Rule::import_decl {
            if let Some(import) = walk_import(pair.clone()) {
                let key = match &import.kind {
                    ImportKind::Simple { path, alias } => {
                        format!("s:{path}:{}", alias.as_deref().unwrap_or(""))
                    }
                    ImportKind::Wildcard { path, alias } => {
                        format!("w:{path}:{}", alias.as_deref().unwrap_or(""))
                    }
                    ImportKind::Named { path, names, level } => {
                        format!("n:{path}:{level}:{names:?}")
                    }
                    ImportKind::Default { path, local } => {
                        format!("d:{path}:{local}")
                    }
                };
                if seen.insert(key) {
                    imports.push(import);
                }
            }
            return;
        }
        for child in pair.into_inner() {
            walk(child, imports, seen);
        }
    }

    let mut imports = Vec::new();
    let mut seen = HashSet::new();
    walk(root.clone(), &mut imports, &mut seen);
    imports
}

fn dotted_ident_expr(path: &str) -> Expression {
    let mut parts = path.split('.');
    let Some(first) = parts.next() else {
        return Expression::ident(path);
    };
    parts.fold(Expression::ident(first), |object, field| {
        Expression::new(ExprKind::Member {
            object: Box::new(object),
            field: field.to_string(),
            null_safe: false,
        })
    })
}

fn kotlin_import_aliases(imports: &[Import]) -> HashMap<String, String> {
    let mut aliases = HashMap::new();
    for import in imports {
        match &import.kind {
            ImportKind::Simple {
                path,
                alias: Some(alias),
            } => {
                if path.rsplit('.').next() != Some(alias.as_str()) {
                    aliases.insert(alias.clone(), path.clone());
                }
            }
            ImportKind::Named { path, names, .. } => {
                for name in names {
                    if let Some(alias) = &name.alias {
                        aliases.insert(alias.clone(), format!("{path}.{}", name.name));
                    }
                }
            }
            _ => {}
        }
    }
    aliases
}

fn imported_leaf(path: &str) -> Option<&str> {
    path.rsplit('.').next()
}

fn imported_leaf_is_value(path: &str) -> bool {
    imported_leaf(path).is_some_and(|leaf| {
        leaf.chars()
            .next()
            .is_some_and(|ch| !ch.is_ascii_uppercase())
    })
}

fn rewrite_imported_value_ident(name: &mut String, aliases: &HashMap<String, String>, scope: &HashSet<String>) {
    if scope.contains(name) {
        return;
    }
    if let Some(path) = aliases.get(name) {
        if imported_leaf_is_value(path) {
            if let Some(leaf) = imported_leaf(path) {
                *name = leaf.to_string();
            }
        }
    }
}

fn imported_type_alias_expr(name: &str, aliases: &HashMap<String, String>, scope: &HashSet<String>) -> Option<Expression> {
    if scope.contains(name) {
        return None;
    }
    let path = aliases.get(name)?;
    if imported_leaf_is_value(path) {
        return None;
    }
    Some(dotted_ident_expr(path))
}

fn kotlin_import_path(path: &str) -> String {
    let Some(leaf) = path.rsplit('.').next() else {
        return path.to_string();
    };
    let java_util = match leaf {
        "ArrayList" | "HashMap" | "HashSet" | "LinkedHashMap" | "LinkedHashSet" => Some(leaf),
        _ => None,
    };
    if path.starts_with("kotlin.collections.") {
        if let Some(leaf) = java_util {
            return format!("java.util.{leaf}");
        }
    }
    if path == "kotlin.text.StringBuilder" {
        return "java.lang.StringBuilder".to_string();
    }
    path.to_string()
}

fn kotlin_import_leaf_is_constant(path: &str) -> bool {
    let Some(leaf) = path.rsplit('.').next() else {
        return false;
    };
    path.starts_with("kotlin.math.")
        && leaf.chars().any(|ch| ch.is_ascii_alphabetic())
        && leaf
            .chars()
            .filter(|ch| ch.is_ascii_alphabetic())
            .all(|ch| ch.is_ascii_uppercase())
}

fn rewrite_import_aliases_in_stmts(stmts: &mut [Statement], aliases: &HashMap<String, String>) {
    let mut scope = HashSet::new();
    for stmt in stmts.iter() {
        collect_declared_names(stmt, &mut scope);
    }
    for stmt in stmts {
        rewrite_import_aliases_in_stmt(stmt, aliases, &mut scope);
    }
}

fn collect_declared_names(stmt: &Statement, names: &mut HashSet<String>) {
    match &stmt.kind {
        StmtKind::ClassDecl { name, .. }
        | StmtKind::InterfaceDecl { name, .. }
        | StmtKind::EnumDecl { name, .. } => {
            names.insert(name.clone());
        }
        StmtKind::FunctionDecl { name, .. } => {
            names.insert(name.clone());
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                collect_binding_names(&decl.pattern, names);
            }
        }
        StmtKind::Block(stmts) => {
            for stmt in stmts {
                collect_declared_names(stmt, names);
            }
        }
        StmtKind::NamespaceDecl { body, .. } => {
            for stmt in body {
                collect_declared_names(stmt, names);
            }
        }
        _ => {}
    }
}

fn collect_binding_names(pattern: &BindingPattern, names: &mut HashSet<String>) {
    match pattern {
        BindingPattern::Ident(name) => {
            names.insert(name.clone());
        }
        BindingPattern::Array(elems) => {
            for elem in elems {
                match elem {
                    ArrayPatternElem::Pattern(pattern, _) => collect_binding_names(pattern, names),
                    ArrayPatternElem::Rest(name) => {
                        names.insert(name.clone());
                    }
                    ArrayPatternElem::Hole => {}
                }
            }
        }
        BindingPattern::Object(props) => {
            for prop in props {
                if prop.is_rest {
                    names.insert(prop.key.clone());
                } else if let Some(pattern) = &prop.value {
                    collect_binding_names(pattern, names);
                } else {
                    names.insert(prop.key.clone());
                }
            }
        }
    }
}

fn rewrite_import_aliases_in_stmt(
    stmt: &mut Statement,
    aliases: &HashMap<String, String>,
    scope: &mut HashSet<String>,
) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            rewrite_import_aliases_in_expr(expr, aliases, scope)
        }
        StmtKind::Throw { expr, cause } => {
            if let Some(expr) = expr {
                rewrite_import_aliases_in_expr(expr, aliases, scope);
            }
            if let Some(cause) = cause {
                rewrite_import_aliases_in_expr(cause, aliases, scope);
            }
        }
        StmtKind::Echo(exprs) => {
            for expr in exprs {
                rewrite_import_aliases_in_expr(expr, aliases, scope);
            }
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_import_aliases_in_expr(init, aliases, scope);
                }
                collect_binding_names(&decl.pattern, scope);
            }
        }
        StmtKind::FunctionDecl { params, body, .. } => {
            let mut fn_scope = scope.clone();
            for param in params {
                fn_scope.insert(param.name.clone());
            }
            for stmt in body.iter() {
                collect_declared_names(stmt, &mut fn_scope);
            }
            rewrite_import_aliases_in_stmts_with_scope(body, aliases, &mut fn_scope);
        }
        StmtKind::ClassDecl { members, .. } => {
            for member in members {
                match member {
                    ClassMember::Field {
                        init: Some(init), ..
                    } => {
                        rewrite_import_aliases_in_expr(init, aliases, scope);
                    }
                    ClassMember::Method(method) => {
                        rewrite_import_aliases_in_stmt(method, aliases, scope);
                    }
                    ClassMember::Constructor { params, body, .. } => {
                        let mut ctor_scope = scope.clone();
                        for param in params {
                            ctor_scope.insert(param.name.clone());
                        }
                        rewrite_import_aliases_in_stmts_with_scope(body, aliases, &mut ctor_scope);
                    }
                    _ => {}
                }
            }
        }
        StmtKind::InterfaceDecl { .. } => {}
        StmtKind::EnumDecl { body_members, .. } => {
            for member in body_members {
                if let ClassMember::Method(method) = member {
                    rewrite_import_aliases_in_stmt(method, aliases, scope);
                }
            }
        }
        StmtKind::Block(stmts) | StmtKind::NamespaceDecl { body: stmts, .. } => {
            let mut inner_scope = scope.clone();
            for stmt in stmts.iter() {
                collect_declared_names(stmt, &mut inner_scope);
            }
            rewrite_import_aliases_in_stmts_with_scope(stmts, aliases, &mut inner_scope);
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_import_aliases_in_expr(cond, aliases, scope);
            rewrite_import_aliases_in_stmts_with_scope(then_body, aliases, &mut scope.clone());
            for (elif_cond, elif_body) in elifs {
                rewrite_import_aliases_in_expr(elif_cond, aliases, scope);
                rewrite_import_aliases_in_stmts_with_scope(elif_body, aliases, &mut scope.clone());
            }
            if let Some(else_body) = else_body {
                rewrite_import_aliases_in_stmts_with_scope(else_body, aliases, &mut scope.clone());
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            rewrite_import_aliases_in_expr(cond, aliases, scope);
            rewrite_import_aliases_in_stmts_with_scope(body, aliases, &mut scope.clone());
            if let Some(else_body) = else_body {
                rewrite_import_aliases_in_stmts_with_scope(else_body, aliases, &mut scope.clone());
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                rewrite_import_aliases_in_stmt(init, aliases, scope);
            }
            if let Some(cond) = cond {
                rewrite_import_aliases_in_expr(cond, aliases, scope);
            }
            if let Some(update) = update {
                rewrite_import_aliases_in_expr(update, aliases, scope);
            }
            rewrite_import_aliases_in_stmts_with_scope(body, aliases, &mut scope.clone());
        }
        StmtKind::ForIn {
            var,
            key,
            iter,
            body,
            else_body,
            ..
        } => {
            rewrite_import_aliases_in_expr(iter, aliases, scope);
            let mut loop_scope = scope.clone();
            loop_scope.insert(var.clone());
            if let Some(key) = key {
                loop_scope.insert(key.clone());
            }
            rewrite_import_aliases_in_stmts_with_scope(body, aliases, &mut loop_scope);
            if let Some(else_body) = else_body {
                rewrite_import_aliases_in_stmts_with_scope(else_body, aliases, &mut scope.clone());
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            rewrite_import_aliases_in_stmts_with_scope(body, aliases, &mut scope.clone());
            for catch in catches {
                let mut catch_scope = scope.clone();
                if let Some(name) = &catch.var_name {
                    catch_scope.insert(name.clone());
                }
                if let Some(name) = &catch.stack_var {
                    catch_scope.insert(name.clone());
                }
                if let Some(when_clause) = &mut catch.when_clause {
                    rewrite_import_aliases_in_expr(when_clause, aliases, &catch_scope);
                }
                rewrite_import_aliases_in_stmts_with_scope(
                    &mut catch.body,
                    aliases,
                    &mut catch_scope,
                );
            }
            if let Some(else_body) = else_body {
                rewrite_import_aliases_in_stmts_with_scope(else_body, aliases, &mut scope.clone());
            }
            if let Some(finally) = finally {
                rewrite_import_aliases_in_stmts_with_scope(finally, aliases, &mut scope.clone());
            }
        }
        _ => {}
    }
}

fn rewrite_import_aliases_in_stmts_with_scope(
    stmts: &mut [Statement],
    aliases: &HashMap<String, String>,
    scope: &mut HashSet<String>,
) {
    for stmt in stmts.iter() {
        collect_declared_names(stmt, scope);
    }
    for stmt in stmts {
        rewrite_import_aliases_in_stmt(stmt, aliases, scope);
    }
}

fn rewrite_import_aliases_in_expr(
    expr: &mut Expression,
    aliases: &HashMap<String, String>,
    scope: &HashSet<String>,
) {
    match &mut expr.kind {
        ExprKind::Ident(name) => {
            rewrite_imported_value_ident(name, aliases, scope);
        }
        ExprKind::Member { object, field, .. } => {
            rewrite_import_aliases_in_expr(object, aliases, scope);
            rewrite_imported_value_ident(field, aliases, scope);
        }
        ExprKind::Call { callee, args, .. } => {
            if let ExprKind::Ident(name) = &mut callee.kind {
                if let Some(replacement) = imported_type_alias_expr(name, aliases, scope) {
                    **callee = replacement;
                } else {
                    rewrite_imported_value_ident(name, aliases, scope);
                }
            } else {
                rewrite_import_aliases_in_expr(callee, aliases, scope);
            }
            for arg in args {
                rewrite_import_aliases_in_expr(&mut arg.value, aliases, scope);
            }
        }
        ExprKind::New { class, args } => {
            if let ExprKind::Ident(name) = &mut class.kind {
                if let Some(replacement) = imported_type_alias_expr(name, aliases, scope) {
                    **class = replacement;
                } else {
                    rewrite_imported_value_ident(name, aliases, scope);
                }
            } else {
                rewrite_import_aliases_in_expr(class, aliases, scope);
            }
            for arg in args {
                rewrite_import_aliases_in_expr(&mut arg.value, aliases, scope);
            }
        }
        ExprKind::Binary { left, right, .. } => {
            rewrite_import_aliases_in_expr(left, aliases, scope);
            rewrite_import_aliases_in_expr(right, aliases, scope);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::Await(expr)
        | ExprKind::Yield(Some(expr))
        | ExprKind::Delete(expr) => rewrite_import_aliases_in_expr(expr, aliases, scope),
        ExprKind::Assign { target, value } => {
            rewrite_import_aliases_in_expr(target, aliases, scope);
            rewrite_import_aliases_in_expr(value, aliases, scope);
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_import_aliases_in_expr(object, aliases, scope);
            rewrite_import_aliases_in_expr(index, aliases, scope);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_import_aliases_in_expr(cond, aliases, scope);
            rewrite_import_aliases_in_expr(then, aliases, scope);
            rewrite_import_aliases_in_expr(else_, aliases, scope);
        }
        ExprKind::NullCoalesce { left, right } => {
            rewrite_import_aliases_in_expr(left, aliases, scope);
            rewrite_import_aliases_in_expr(right, aliases, scope);
        }
        ExprKind::Array(elems) => {
            for elem in elems {
                if let Some(key) = &mut elem.key {
                    rewrite_import_aliases_in_expr(key, aliases, scope);
                }
                rewrite_import_aliases_in_expr(&mut elem.value, aliases, scope);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                if let ObjectProperty::KeyValue { key, value } = prop {
                    rewrite_import_aliases_in_expr(key, aliases, scope);
                    rewrite_import_aliases_in_expr(value, aliases, scope);
                }
            }
        }
        ExprKind::Tuple(items) => {
            for item in items {
                rewrite_import_aliases_in_expr(item, aliases, scope);
            }
        }
        ExprKind::Range { start, end, .. } => {
            rewrite_import_aliases_in_expr(start, aliases, scope);
            rewrite_import_aliases_in_expr(end, aliases, scope);
        }
        ExprKind::Lambda { params, body, .. } => {
            let mut lambda_scope = scope.clone();
            for param in params {
                lambda_scope.insert(param.name.clone());
            }
            match body {
                LambdaBody::Expr(expr) => {
                    rewrite_import_aliases_in_expr(expr, aliases, &lambda_scope);
                }
                LambdaBody::Block(stmts) => {
                    rewrite_import_aliases_in_stmts_with_scope(stmts, aliases, &mut lambda_scope);
                }
            }
        }
        _ => {}
    }
    if let Some(replacement) = post_alias_kotlin_lowering(expr) {
        *expr = replacement;
    }
}

fn post_alias_kotlin_lowering(expr: &Expression) -> Option<Expression> {
    match &expr.kind {
        ExprKind::Member { object, field, .. } if field == "absoluteValue" => {
            Some(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("absoluteValue")),
                args: vec![Argument::positional(*object.clone())],
                optional: false,
            }))
        }
        ExprKind::Call { callee, args, .. } => match &callee.kind {
            ExprKind::Ident(name) if name == "joinToString" && !args.is_empty() => {
                let separator = args
                    .get(1)
                    .map(|arg| arg.value.clone())
                    .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Str(", ".into()))));
                Some(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__coll_join")),
                    args: vec![
                        Argument::positional(args[0].value.clone()),
                        Argument::positional(separator),
                    ],
                    optional: false,
                }))
            }
            ExprKind::Member { object, field, .. } if field == "joinToString" => {
                let separator = args
                    .first()
                    .map(|arg| arg.value.clone())
                    .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Str(", ".into()))));
                Some(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__coll_join")),
                    args: vec![
                        Argument::positional(*object.clone()),
                        Argument::positional(separator),
                    ],
                    optional: false,
                }))
            }
            ExprKind::Member { object, field, .. } if field == "sortedBy" && args.len() == 1 => {
                Some(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Member {
                        object: object.clone(),
                        field: "sort".to_string(),
                        null_safe: false,
                    })),
                    args: args.clone(),
                    optional: false,
                }))
            }
            ExprKind::Member { object, field, .. } if field == "let" && args.len() == 1 => {
                Some(Expression::new(ExprKind::Call {
                    callee: Box::new(args[0].value.clone()),
                    args: vec![Argument::positional(*object.clone())],
                    optional: false,
                }))
            }
            _ => None,
        },
        _ => None,
    }
}

#[derive(Debug, Clone, Default)]
struct KotlinOperatorInfo {
    returns: HashMap<String, Option<String>>,
}

impl KotlinOperatorInfo {
    fn has(&self, name: &str) -> bool {
        self.returns.contains_key(name)
    }

    fn return_type(&self, name: &str) -> Option<String> {
        self.returns.get(name).and_then(Clone::clone)
    }
}

type KotlinOperatorTable = HashMap<String, KotlinOperatorInfo>;
type KotlinLocalTypes = HashMap<String, String>;

fn normalize_kotlin_operator_calls(stmts: &mut [Statement]) {
    let operators = collect_kotlin_operator_table(stmts);
    let mut locals = KotlinLocalTypes::new();
    normalize_kotlin_operator_stmts(stmts, &operators, &mut locals);
}

fn collect_kotlin_operator_table(stmts: &[Statement]) -> KotlinOperatorTable {
    fn visit_stmt(stmt: &Statement, out: &mut KotlinOperatorTable) {
        match &stmt.kind {
            StmtKind::ClassDecl { name, members, .. } => {
                for member in members {
                    if let ClassMember::Method(method) = member {
                        if let StmtKind::FunctionDecl {
                            name: method_name,
                            return_type,
                            ..
                        } = &method.kind
                        {
                            if let Some(op_name) = method_name.strip_prefix("operator ") {
                                out.entry(name.clone())
                                    .or_default()
                                    .returns
                                    .insert(op_name.to_string(), return_type.clone());
                            }
                        }
                    }
                }
                for member in members {
                    match member {
                        ClassMember::Constructor { body, .. } => {
                            for stmt in body {
                                visit_stmt(stmt, out);
                            }
                        }
                        ClassMember::Field { .. }
                        | ClassMember::Method(_)
                        | ClassMember::Property { .. }
                        | ClassMember::Event { .. }
                        | ClassMember::Const { .. } => {}
                        ClassMember::NestedType(class_stmt) => visit_stmt(class_stmt, out),
                        ClassMember::Augment(_) => {}
                    }
                }
            }
            StmtKind::NamespaceDecl { body, .. } | StmtKind::Block(body) => {
                for stmt in body {
                    visit_stmt(stmt, out);
                }
            }
            _ => {}
        }
    }

    let mut out = KotlinOperatorTable::new();
    for stmt in stmts {
        visit_stmt(stmt, &mut out);
    }
    out
}

fn normalize_kotlin_operator_stmts(
    stmts: &mut [Statement],
    operators: &KotlinOperatorTable,
    locals: &mut KotlinLocalTypes,
) {
    for stmt in stmts {
        normalize_kotlin_operator_stmt(stmt, operators, locals);
    }
}

fn normalize_kotlin_operator_stmt(
    stmt: &mut Statement,
    operators: &KotlinOperatorTable,
    locals: &mut KotlinLocalTypes,
) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            normalize_kotlin_operator_expr(expr, operators, locals);
        }
        StmtKind::Throw { expr, cause } => {
            if let Some(expr) = expr {
                normalize_kotlin_operator_expr(expr, operators, locals);
            }
            if let Some(cause) = cause {
                normalize_kotlin_operator_expr(cause, operators, locals);
            }
        }
        StmtKind::Echo(exprs) => {
            for expr in exprs {
                normalize_kotlin_operator_expr(expr, operators, locals);
            }
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    normalize_kotlin_operator_expr(init, operators, locals);
                }
                if let BindingPattern::Ident(name) = &decl.pattern {
                    let inferred = decl
                        .type_hint
                        .clone()
                        .or_else(|| decl.init.as_ref().and_then(|e| kotlin_expr_type(e, locals, operators)));
                    if let Some(ty) = inferred {
                        locals.insert(name.clone(), ty);
                    }
                }
            }
        }
        StmtKind::FunctionDecl { params, body, .. } => {
            let mut fn_locals = locals.clone();
            for param in params {
                if let Some(ty) = &param.type_hint {
                    fn_locals.insert(param.name.clone(), ty.clone());
                }
            }
            normalize_kotlin_operator_stmts(body, operators, &mut fn_locals);
        }
        StmtKind::ClassDecl { members, .. } => {
            for member in members {
                match member {
                    ClassMember::Field {
                        init: Some(init), ..
                    } => normalize_kotlin_operator_expr(init, operators, locals),
                    ClassMember::Method(method) => {
                        normalize_kotlin_operator_stmt(method, operators, locals);
                    }
                    ClassMember::Constructor { params, body, .. } => {
                        let mut ctor_locals = locals.clone();
                        for param in params {
                            if let Some(ty) = &param.type_hint {
                                ctor_locals.insert(param.name.clone(), ty.clone());
                            }
                        }
                        normalize_kotlin_operator_stmts(body, operators, &mut ctor_locals);
                    }
                    ClassMember::NestedType(class_stmt) => {
                        normalize_kotlin_operator_stmt(class_stmt, operators, locals);
                    }
                    _ => {}
                }
            }
        }
        StmtKind::Block(body) | StmtKind::NamespaceDecl { body, .. } => {
            let mut block_locals = locals.clone();
            normalize_kotlin_operator_stmts(body, operators, &mut block_locals);
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            normalize_kotlin_operator_expr(cond, operators, locals);
            normalize_kotlin_operator_stmts(then_body, operators, &mut locals.clone());
            for (elif_cond, elif_body) in elifs {
                normalize_kotlin_operator_expr(elif_cond, operators, locals);
                normalize_kotlin_operator_stmts(elif_body, operators, &mut locals.clone());
            }
            if let Some(else_body) = else_body {
                normalize_kotlin_operator_stmts(else_body, operators, &mut locals.clone());
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            normalize_kotlin_operator_expr(cond, operators, locals);
            normalize_kotlin_operator_stmts(body, operators, &mut locals.clone());
            if let Some(else_body) = else_body {
                normalize_kotlin_operator_stmts(else_body, operators, &mut locals.clone());
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            let mut for_locals = locals.clone();
            if let Some(init) = init {
                normalize_kotlin_operator_stmt(init, operators, &mut for_locals);
            }
            if let Some(cond) = cond {
                normalize_kotlin_operator_expr(cond, operators, &for_locals);
            }
            if let Some(update) = update {
                normalize_kotlin_operator_expr(update, operators, &for_locals);
            }
            normalize_kotlin_operator_stmts(body, operators, &mut for_locals);
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            normalize_kotlin_operator_expr(iter, operators, locals);
            normalize_kotlin_operator_stmts(body, operators, &mut locals.clone());
            if let Some(else_body) = else_body {
                normalize_kotlin_operator_stmts(else_body, operators, &mut locals.clone());
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            normalize_kotlin_operator_stmts(body, operators, &mut locals.clone());
            for catch in catches {
                if let Some(when_clause) = &mut catch.when_clause {
                    normalize_kotlin_operator_expr(when_clause, operators, locals);
                }
                normalize_kotlin_operator_stmts(&mut catch.body, operators, &mut locals.clone());
            }
            if let Some(else_body) = else_body {
                normalize_kotlin_operator_stmts(else_body, operators, &mut locals.clone());
            }
            if let Some(finally) = finally {
                normalize_kotlin_operator_stmts(finally, operators, &mut locals.clone());
            }
        }
        _ => {}
    }
}

fn normalize_kotlin_operator_expr(
    expr: &mut Expression,
    operators: &KotlinOperatorTable,
    locals: &KotlinLocalTypes,
) {
    match &mut expr.kind {
        ExprKind::Binary { left, right, .. } => {
            normalize_kotlin_operator_expr(left, operators, locals);
            normalize_kotlin_operator_expr(right, operators, locals);
        }
        ExprKind::Unary { expr: inner, .. }
        | ExprKind::Cast { expr: inner, .. }
        | ExprKind::Await(inner)
        | ExprKind::Yield(Some(inner))
        | ExprKind::Delete(inner) => normalize_kotlin_operator_expr(inner, operators, locals),
        ExprKind::Assign { target, value } => {
            normalize_kotlin_operator_expr(target, operators, locals);
            normalize_kotlin_operator_expr(value, operators, locals);
        }
        ExprKind::Call { callee, args, .. } => {
            normalize_kotlin_operator_expr(callee, operators, locals);
            for arg in args {
                normalize_kotlin_operator_expr(&mut arg.value, operators, locals);
            }
        }
        ExprKind::New { class, args } => {
            normalize_kotlin_operator_expr(class, operators, locals);
            for arg in args {
                normalize_kotlin_operator_expr(&mut arg.value, operators, locals);
            }
        }
        ExprKind::Member { object, .. } => normalize_kotlin_operator_expr(object, operators, locals),
        ExprKind::Index { object, index, .. } => {
            normalize_kotlin_operator_expr(object, operators, locals);
            normalize_kotlin_operator_expr(index, operators, locals);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            normalize_kotlin_operator_expr(cond, operators, locals);
            normalize_kotlin_operator_expr(then, operators, locals);
            normalize_kotlin_operator_expr(else_, operators, locals);
        }
        ExprKind::NullCoalesce { left, right } => {
            normalize_kotlin_operator_expr(left, operators, locals);
            normalize_kotlin_operator_expr(right, operators, locals);
        }
        ExprKind::Array(elems) => {
            for elem in elems {
                if let Some(key) = &mut elem.key {
                    normalize_kotlin_operator_expr(key, operators, locals);
                }
                normalize_kotlin_operator_expr(&mut elem.value, operators, locals);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                if let ObjectProperty::KeyValue { key, value } = prop {
                    normalize_kotlin_operator_expr(key, operators, locals);
                    normalize_kotlin_operator_expr(value, operators, locals);
                }
            }
        }
        ExprKind::Tuple(items) | ExprKind::Sequence(items) => {
            for item in items {
                normalize_kotlin_operator_expr(item, operators, locals);
            }
        }
        ExprKind::Range { start, end, .. } => {
            normalize_kotlin_operator_expr(start, operators, locals);
            normalize_kotlin_operator_expr(end, operators, locals);
        }
        ExprKind::Lambda { params, body, .. } => {
            let mut lambda_locals = locals.clone();
            for param in params {
                if let Some(ty) = &param.type_hint {
                    lambda_locals.insert(param.name.clone(), ty.clone());
                }
            }
            match body {
                LambdaBody::Expr(expr) => normalize_kotlin_operator_expr(expr, operators, &lambda_locals),
                LambdaBody::Block(stmts) => {
                    normalize_kotlin_operator_stmts(stmts, operators, &mut lambda_locals);
                }
            }
        }
        _ => {}
    }

    if let Some(replacement) = kotlin_operator_rewrite(expr, operators, locals) {
        *expr = replacement;
    }
}

fn kotlin_operator_rewrite(
    expr: &Expression,
    operators: &KotlinOperatorTable,
    locals: &KotlinLocalTypes,
) -> Option<Expression> {
    match &expr.kind {
        ExprKind::Binary { op, left, right } => {
            if *op == BinOp::In {
                if matches!(right.kind, ExprKind::Range { .. })
                    || kotlin_expr_type(right, locals, operators).as_deref() == Some("Range")
                {
                    return Some(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__coll_contains")),
                        args: vec![
                            Argument::positional((**right).clone()),
                            Argument::positional((**left).clone()),
                        ],
                        optional: false,
                    }));
                }
                let ty = kotlin_expr_type(right, locals, operators)?;
                let method = crate::protocol::binary_operator_method(*op)?;
                if operators.get(&ty).is_some_and(|info| info.has(method)) {
                    return Some(kotlin_operator_call((**right).clone(), method, vec![
                        (**left).clone(),
                    ]));
                }
                return None;
            }

            let method = crate::protocol::binary_operator_method(*op)?;
            let ty = kotlin_expr_type(left, locals, operators)?;
            if !operators.get(&ty).is_some_and(|info| info.has(method)) {
                return None;
            }
            let call = kotlin_operator_call((**left).clone(), method, vec![(**right).clone()]);
            match op {
                BinOp::Lt => Some(kotlin_compare_zero_call("__kt_cmp_lt0", call)),
                BinOp::Gt => Some(kotlin_compare_zero_call("__kt_cmp_gt0", call)),
                BinOp::LtEq => Some(kotlin_compare_zero_call("__kt_cmp_le0", call)),
                BinOp::GtEq => Some(kotlin_compare_zero_call("__kt_cmp_ge0", call)),
                _ => Some(call),
            }
        }
        ExprKind::Assign { target, value } => {
            let ExprKind::Binary { op, left, right } = &value.kind else {
                return None;
            };
            if !kotlin_same_simple_expr(target, left) {
                return None;
            }
            let ty = kotlin_expr_type(target, locals, operators)?;
            let info = operators.get(&ty)?;
            if let Some(method) = crate::protocol::compound_operator_method(*op) {
                if info.has(method) {
                    return Some(kotlin_operator_call(
                        (**target).clone(),
                        method,
                        vec![(**right).clone()],
                    ));
                }
            }

            if !matches!(right.kind, ExprKind::Lit(Literal::Int(1))) {
                return None;
            }
            let method = crate::protocol::step_operator_method(*op)?;
            info.has(method).then(|| Expression::new(ExprKind::Assign {
                target: Box::new((**target).clone()),
                value: Box::new(kotlin_operator_call((**target).clone(), method, Vec::new())),
            }))
        }
        ExprKind::Unary { op, expr: inner } => {
            let method = crate::protocol::unary_operator_method(*op)?;
            let ty = kotlin_expr_type(inner, locals, operators)?;
            operators
                .get(&ty)
                .is_some_and(|info| info.has(method))
                .then(|| kotlin_operator_call((**inner).clone(), method, Vec::new()))
        }
        _ => None,
    }
}

fn kotlin_same_simple_expr(a: &Expression, b: &Expression) -> bool {
    match (&a.kind, &b.kind) {
        (ExprKind::Ident(a), ExprKind::Ident(b)) => a == b,
        (
            ExprKind::Member {
                object: ao,
                field: af,
                null_safe: ans,
            },
            ExprKind::Member {
                object: bo,
                field: bf,
                null_safe: bns,
            },
        ) => af == bf && ans == bns && kotlin_same_simple_expr(ao, bo),
        (
            ExprKind::Index {
                object: ao,
                index: ai,
                null_safe: ans,
            },
            ExprKind::Index {
                object: bo,
                index: bi,
                null_safe: bns,
            },
        ) => ans == bns && kotlin_same_simple_expr(ao, bo) && kotlin_same_simple_expr(ai, bi),
        (ExprKind::Lit(Literal::Int(a)), ExprKind::Lit(Literal::Int(b))) => a == b,
        (ExprKind::Lit(Literal::Str(a)), ExprKind::Lit(Literal::Str(b))) => a == b,
        _ => false,
    }
}

fn kotlin_operator_call(receiver: Expression, method: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(receiver),
            field: method.to_string(),
            null_safe: false,
        })),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

fn kotlin_compare_zero_call(helper: &str, value: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(helper)),
        args: vec![Argument::positional(value)],
        optional: false,
    })
}

fn kotlin_expr_type(
    expr: &Expression,
    locals: &KotlinLocalTypes,
    operators: &KotlinOperatorTable,
) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => locals.get(name).cloned(),
        ExprKind::New { class, .. } => match &class.kind {
            ExprKind::Ident(name) => Some(name.clone()),
            ExprKind::Member { field, .. } => Some(field.clone()),
            _ => None,
        },
        ExprKind::Call { callee, .. } => {
            let ExprKind::Member { object, field, .. } = &callee.kind else {
                return None;
            };
            let receiver_ty = kotlin_expr_type(object, locals, operators)?;
            operators
                .get(&receiver_ty)
                .and_then(|info| info.return_type(field))
        }
        ExprKind::Binary { op, left, right } => {
            if matches!(
                op,
                BinOp::In
                    | BinOp::Lt
                    | BinOp::Gt
                    | BinOp::LtEq
                    | BinOp::GtEq
                    | BinOp::Eq
                    | BinOp::NotEq
                    | BinOp::StrictEq
                    | BinOp::StrictNotEq
            ) {
                return Some("Boolean".to_string());
            }
            let method = crate::protocol::binary_operator_method(*op)?;
            let receiver_ty = kotlin_expr_type(left, locals, operators)?;
            operators
                .get(&receiver_ty)
                .and_then(|info| info.return_type(method))
                .or_else(|| kotlin_expr_type(right, locals, operators))
        }
        ExprKind::Unary { op, expr } => {
            if matches!(op, UnaryOp::Not) {
                return Some("Boolean".to_string());
            }
            let method = crate::protocol::unary_operator_method(*op)?;
            let receiver_ty = kotlin_expr_type(expr, locals, operators)?;
            operators
                .get(&receiver_ty)
                .and_then(|info| info.return_type(method))
        }
        ExprKind::Range { .. } => Some("Range".to_string()),
        _ => None,
    }
}

fn walk_typealias(_pair: Pair<Rule>) -> Option<Statement> {
    Some(Statement::new(StmtKind::Empty))
}

fn walk_annotation(pair: Pair<Rule>) -> Expression {
    let mut type_name = String::new();
    let mut args = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::type_ref => type_name = inner.as_str().to_string(),
            Rule::arg_list => {
                for arg_p in inner.into_inner() {
                    let mut arg_expr = None;
                    let mut arg_name = None;
                    for sub in arg_p.into_inner() {
                        match sub.as_rule() {
                            Rule::identifier => arg_name = Some(sub.as_str().to_string()),
                            Rule::expr => arg_expr = Some(walk_expr(sub)),
                            _ => {}
                        }
                    }
                    if let Some(ae) = arg_expr {
                        args.push(Argument {
                            value: ae,
                            name: arg_name,
                            by_ref: false,
                            spread: false,
                        });
                    }
                }
            }
            _ => {}
        }
    }

    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(&type_name)),
        args,
        optional: false,
    })
}

/// True when `expr` is a bare dotted chain of identifiers (`java.util`), which
/// is what distinguishes a package-qualified type name from member access on a
/// value.
fn is_ident_chain(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Ident(_) => true,
        ExprKind::Member { object, .. } => is_ident_chain(object),
        _ => false,
    }
}

fn callable_ref_lambda(target: Expression) -> Expression {
    let arg_name = "__kt_ref_arg".to_string();
    Expression::new(ExprKind::Lambda {
        params: vec![Param {
            name: arg_name.clone(),
            type_hint: None,
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        }],
        body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Call {
            callee: Box::new(target),
            args: vec![Argument::positional(Expression::ident(&arg_name))],
            optional: false,
        }))),
        is_async: false,
        captures: Vec::new(),
    })
}

fn walk_callable_ref(pair: Pair<Rule>) -> Expression {
    let mut qualifier = None;
    let mut name = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::dotted_name => qualifier = Some(inner.as_str().to_string()),
            Rule::identifier => name = Some(inner.as_str().to_string()),
            _ => {}
        }
    }

    let Some(name) = name else {
        return Expression::null();
    };
    let target = qualifier
        .map(|qualifier| dotted_ident_expr(&format!("{qualifier}.{name}")))
        .unwrap_or_else(|| Expression::ident(&name));
    callable_ref_lambda(target)
}

fn walk_import(pair: Pair<Rule>) -> Option<Import> {
    let mut path = String::new();
    let mut alias = None;
    let mut is_wildcard = false;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::dotted_name => path = inner.as_str().to_string(),
            Rule::identifier => alias = Some(inner.as_str().to_string()),
            _ => {
                if inner.as_str() == ".*" || inner.as_str() == "*" {
                    is_wildcard = true;
                }
            }
        }
    }

    path = kotlin_import_path(&path);
    let kind = if is_wildcard {
        ImportKind::Wildcard { path, alias }
    } else if kotlin_import_leaf_is_constant(&path) {
        let (module, name) = path.rsplit_once('.').unwrap_or(("", path.as_str()));
        ImportKind::Named {
            path: module.to_string(),
            names: vec![ImportName {
                name: name.to_string(),
                alias,
            }],
            level: 0,
        }
    } else {
        // Kotlin's import binds the SIMPLE NAME (`import java.time.Instant`
        // makes `Instant` mean `java.time.Instant`), so an import with no
        // `as` clause still carries an alias — its last segment. Without it
        // the import was inert: `Instant.parse(…)` resolved to nothing and
        // trapped "undefined is not callable". Java's `walk_import` has done
        // this all along; it is what populates `source_type_aliases`.
        let alias = alias.or_else(|| path.rsplit('.').next().map(str::to_string));
        ImportKind::Simple { path, alias }
    };

    Some(Import {
        kind,
        span: Span::default(),
    })
}

fn walk_statement(pair: Pair<Rule>) -> Option<Statement> {
    let mut label_name = None;

    let inner_pair = if pair.as_rule() == Rule::statement {
        let mut inner_iter = pair.into_inner();
        let first = inner_iter.next()?;
        if first.as_rule() == Rule::label_decl {
            label_name = Some(first.as_str().trim_end_matches('@').to_string());
            inner_iter.next()?
        } else {
            first
        }
    } else {
        pair
    };

    let stmt = match inner_pair.as_rule() {
        Rule::import_decl => Some(Statement::new(StmtKind::Empty)),
        Rule::typealias_decl => walk_typealias(inner_pair),
        Rule::interface_decl => walk_interface_decl(inner_pair),
        Rule::enum_decl => walk_enum_decl(inner_pair),
        Rule::destructuring_decl => walk_destructuring_decl(inner_pair),
        Rule::function_decl => walk_function_decl(inner_pair),
        Rule::var_decl => {
            // An extension property is not a variable at all.
            if inner_pair
                .clone()
                .into_inner()
                .any(|p| p.as_rule() == Rule::receiver_prefix)
            {
                walk_extension_property(inner_pair)
            } else {
                walk_var_decl(inner_pair)
            }
        }
        Rule::class_decl => walk_class_decl(inner_pair),
        Rule::object_decl => walk_object_decl(inner_pair),
        Rule::if_expr => walk_if_stmt(inner_pair),
        Rule::when_expr => walk_when_stmt(inner_pair),
        Rule::try_expr => walk_try_stmt(inner_pair),
        Rule::throw_stmt => {
            let expr = inner_pair
                .into_inner()
                .find(|p| p.as_rule() == Rule::expr)
                .map(walk_expr);
            Some(Statement::new(StmtKind::Throw { expr, cause: None }))
        }
        Rule::for_stmt => walk_for_stmt(inner_pair),
        Rule::while_stmt => walk_while_stmt(inner_pair),
        Rule::do_while_stmt => walk_do_while_stmt(inner_pair),
        Rule::return_stmt => {
            let mut ret_expr = None;
            for rsub in inner_pair.into_inner() {
                if rsub.as_rule() == Rule::expr {
                    ret_expr = Some(walk_expr(rsub));
                }
            }
            Some(Statement::new(StmtKind::Return(ret_expr)))
        }
        Rule::break_stmt => {
            let mut lbl = None;
            for bsub in inner_pair.into_inner() {
                if bsub.as_rule() == Rule::identifier {
                    lbl = Some(bsub.as_str().to_string());
                }
            }
            let target = lbl.map(BreakTarget::Label).unwrap_or(BreakTarget::Implicit);
            Some(Statement::new(StmtKind::Break(target)))
        }
        Rule::continue_stmt => {
            let mut lbl = None;
            for csub in inner_pair.into_inner() {
                if csub.as_rule() == Rule::identifier {
                    lbl = Some(csub.as_str().to_string());
                }
            }
            let target = lbl
                .map(ContinueTarget::Label)
                .unwrap_or(ContinueTarget::Implicit);
            Some(Statement::new(StmtKind::Continue(target)))
        }
        Rule::expr_stmt => {
            let expr_pair = inner_pair.into_inner().next()?;
            let expr = walk_expr(expr_pair);
            Some(repeat_to_for_in(&expr).unwrap_or_else(|| Statement::new(StmtKind::Expr(expr))))
        }
        Rule::expr => {
            let expr = walk_expr(inner_pair);
            Some(Statement::new(StmtKind::Expr(expr)))
        }
        _ => None,
    };

    match (stmt, label_name) {
        (Some(s), Some(lbl)) => Some(Statement::new(StmtKind::Labeled {
            label: lbl,
            body: Box::new(s),
        })),
        (other, _) => other,
    }
}

/// `repeat(n) { … }` -> the `for` loop it stands for.
///
/// Kotlin spells this control structure as a function, but it IS a loop: the
/// lambda runs `n` times and receives the 0-based index. Desugaring it here
/// rather than adapting it to a call is what puts `break` and `continue` inside
/// it on the shared loop machinery, and what makes a label on it mean what a
/// label on any other Kotlin loop means.
fn repeat_to_for_in(expr: &Expression) -> Option<Statement> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if !matches!(&callee.kind, ExprKind::Ident(n) if n == "repeat") || args.len() != 2 {
        return None;
    }
    let ExprKind::Lambda { params, body, .. } = &args[1].value.kind else {
        return None;
    };
    let var = params
        .first()
        .map(|p| p.name.clone())
        .unwrap_or_else(|| "it".to_string());
    let body = match body {
        LambdaBody::Block(stmts) => stmts.clone(),
        LambdaBody::Expr(e) => vec![Statement::new(StmtKind::Expr((**e).clone()))],
    };
    Some(Statement::new(StmtKind::ForIn {
        var,
        key: None,
        iter: Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__kt_step_asc")),
            args: vec![
                Argument::positional(Expression::int(0)),
                Argument::positional(args[0].value.clone()),
                Argument::positional(Expression::int(1)),
            ],
            optional: false,
        }),
        body,
        of: true,
        else_body: None,
        is_async: false,
    }))
}

fn walk_interface_decl(pair: Pair<Rule>) -> Option<Statement> {
    let mut name = String::new();
    let mut parents = Vec::new();
    let mut members = Vec::new();
    let mut decorators = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::annotation => decorators.push(walk_annotation(inner)),
            Rule::identifier => {
                if name.is_empty() {
                    name = inner.as_str().to_string();
                }
            }
            Rule::inheritance_list => {
                for spec in inner.into_inner() {
                    if spec.as_rule() == Rule::inheritance_specifier {
                        for sub in spec.into_inner() {
                            if sub.as_rule() == Rule::type_ref {
                                let base = sub.as_str().trim().to_string();
                                // `interface B : A` — B carries A's defaults, so
                                // a class implementing only B still gets them.
                                // The fold order resolves the chain, so B is
                                // augmented before any class that implements it.
                                members.push(ClassMember::Augment(AugmentDecl {
                                    from: base.clone(),
                                    via_field: None,
                                    adjustments: vec![],
                                }));
                                parents.push(base);
                            }
                        }
                    }
                }
            }
            Rule::class_body => {
                for member_pair in inner.into_inner() {
                    if member_pair.as_rule() == Rule::class_member {
                        if let Some(inner_member) = member_pair.into_inner().next() {
                            match inner_member.as_rule() {
                                Rule::function_decl => {
                                    if let Some(mut stmt) = walk_function_decl(inner_member) {
                                        // A Kotlin interface method with no
                                        // block is abstract; one WITH a block is
                                        // a default implementation, and the body
                                        // has to survive — `InterfaceMember`
                                        // had nowhere to put it, so every
                                        // default method was silently emptied.
                                        if let StmtKind::FunctionDecl {
                                            body, modifiers, ..
                                        } = &mut stmt.kind
                                        {
                                            modifiers.is_abstract = body.is_empty();
                                        }
                                        members.push(ClassMember::Method(Box::new(stmt)));
                                    }
                                }
                                Rule::var_decl => {
                                    let prop = walk_class_property(inner_member.clone());
                                    if !prop.is_empty() {
                                        members.extend(prop);
                                        continue;
                                    }
                                    if let Some(stmt) = walk_var_decl(inner_member) {
                                        if let StmtKind::VarDecl { declarations, kind } = stmt.kind
                                        {
                                            for decl in declarations {
                                                if let BindingPattern::Ident(pname) = decl.pattern {
                                                    members.push(ClassMember::Field {
                                                        name: pname,
                                                        type_hint: decl.type_hint,
                                                        init: decl.init,
                                                        modifiers: Modifiers {
                                                            visibility: Visibility::Public,
                                                            is_readonly: kind == VarDeclKind::Const,
                                                            ..Default::default()
                                                        },
                                                        with_events: false,
                                                        array_bounds: None,
                                                    });
                                                }
                                            }
                                        }
                                    }
                                }
                                _ => {}
                            }
                        }
                    }
                }
            }
            _ => {}
        }
    }

    // An interface is a CLASS DECLARATION whose `declared_kind` says
    // `Interface` — flexclassplan §0.1's one class model, and the shape
    // `ClassKind` exists for. As a `StmtKind::InterfaceDecl` it never entered
    // `normalized_classes`, so `class W(d: I) : I by d` could not find `I`'s
    // members to promote and delegation resolved to nothing.
    Some(Statement::new(StmtKind::ClassDecl {
        name,
        parents: Vec::new(),
        // A Kotlin interface's supertypes are other interfaces, never a
        // superclass — so they are the interface list, not `parents`.
        interfaces: parents,
        members,
        modifiers: ClassModifiers {
            is_abstract: true,
            kind: ClassKind::Interface,
            ..Default::default()
        },
        decorators,
    }))
}

fn walk_enum_decl(pair: Pair<Rule>) -> Option<Statement> {
    let mut name = String::new();
    let mut members = Vec::new();
    let mut body_members = Vec::new();
    let mut decorators = Vec::new();
    let mut entry_idx = 0i64;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::annotation => decorators.push(walk_annotation(inner)),
            Rule::identifier => {
                if name.is_empty() {
                    name = inner.as_str().to_string();
                }
            }
            Rule::enum_entry => {
                let mut em_name = String::new();
                let mut ctor_args = Vec::new();
                for esub in inner.into_inner() {
                    match esub.as_rule() {
                        Rule::identifier => em_name = esub.as_str().to_string(),
                        Rule::arg_list => {
                            for arg_p in esub.into_inner() {
                                for e in arg_p.into_inner() {
                                    if e.as_rule() == Rule::expr {
                                        ctor_args.push(walk_expr(e));
                                    }
                                }
                            }
                        }
                        _ => {}
                    }
                }
                if !em_name.is_empty() {
                    let val_expr = if let Some(first_arg) = ctor_args.first() {
                        Some(first_arg.clone())
                    } else {
                        Some(Expression::int(entry_idx))
                    };
                    entry_idx += 1;
                    members.push(EnumMember {
                        name: em_name,
                        value: val_expr,
                        constructor_args: ctor_args,
                    });
                }
            }
            Rule::class_member => {
                if let Some(inner_member) = inner.into_inner().next() {
                    if inner_member.as_rule() == Rule::function_decl {
                        if let Some(stmt) = walk_function_decl(inner_member) {
                            body_members.push(ClassMember::Method(Box::new(stmt)));
                        }
                    }
                }
            }
            _ => {}
        }
    }

    Some(Statement::new(StmtKind::EnumDecl {
        name,
        members,
        visibility: Visibility::Public,
        is_flags: false,
        backing_type: None,
        interfaces: vec![],
        body_members,
        decorators,
    }))
}

fn walk_destructuring_decl(pair: Pair<Rule>) -> Option<Statement> {
    let mut is_readonly = false;
    let mut names = Vec::new();
    let mut init = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::val_kw => is_readonly = true,
            Rule::var_kw => is_readonly = false,
            Rule::destructuring_target => {
                for target_inner in inner.into_inner() {
                    if target_inner.as_rule() == Rule::identifier {
                        names.push(target_inner.as_str().to_string());
                    }
                }
            }
            Rule::identifier => names.push(inner.as_str().to_string()),
            Rule::expr => init = Some(walk_expr(inner)),
            _ => {}
        }
    }

    if let Some(init_expr) = init {
        let tmp_name = gen_tmp_name();
        let decl_kind = if is_readonly {
            VarDeclKind::Const
        } else {
            VarDeclKind::Var
        };

        let mut stmts = vec![Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(tmp_name.clone()),
                type_hint: None,
                init: Some(init_expr),
                array_bounds: None,
                with_events: false,
            }],
            kind: decl_kind.clone(),
        })];

        for (idx, name) in names.into_iter().enumerate() {
            let read_expr = Expression::new(ExprKind::Index {
                object: Box::new(Expression::ident(&tmp_name)),
                index: Box::new(Expression::int(idx as i64)),
                null_safe: false,
            });
            stmts.push(Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(name),
                    type_hint: None,
                    init: Some(read_expr),
                    array_bounds: None,
                    with_events: false,
                }],
                kind: decl_kind.clone(),
            }));
        }

        Some(Statement::new(StmtKind::Block(stmts)))
    } else {
        let elems = names
            .into_iter()
            .map(|n| ArrayPatternElem::Pattern(BindingPattern::Ident(n), None))
            .collect();
        Some(Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Array(elems),
                type_hint: None,
                init: None,
                array_bounds: None,
                with_events: false,
            }],
            kind: if is_readonly {
                VarDeclKind::Const
            } else {
                VarDeclKind::Var
            },
        }))
    }
}

fn walk_try_stmt(pair: Pair<Rule>) -> Option<Statement> {
    let mut body = Vec::new();
    let mut catches = Vec::new();
    let mut finally = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::block => {
                if body.is_empty() {
                    body = walk_block_statements(inner);
                }
            }
            Rule::catch_clause => {
                let mut param_name = "e".to_string();
                let mut type_hint = None;
                let mut catch_block_stmts = Vec::new();
                for csub in inner.into_inner() {
                    match csub.as_rule() {
                        Rule::identifier => param_name = csub.as_str().to_string(),
                        Rule::type_ref => type_hint = Some(type_hint_text(csub.as_str())),
                        Rule::block => catch_block_stmts = walk_block_statements(csub),
                        _ => {}
                    }
                }
                let types = match type_hint.as_deref() {
                    Some("Exception") | Some("Throwable") | None => vec![],
                    Some(t) => vec![t.to_string()],
                };
                catches.push(CatchClause {
                    types,
                    var_name: Some(param_name),
                    stack_var: None,
                    body: catch_block_stmts,
                    when_clause: None,
                });
            }
            Rule::finally_clause => {
                for fsub in inner.into_inner() {
                    if fsub.as_rule() == Rule::block {
                        finally = Some(walk_block_statements(fsub));
                    }
                }
            }
            _ => {}
        }
    }

    Some(Statement::new(StmtKind::Try {
        body,
        catches,
        else_body: None,
        finally,
    }))
}

fn walk_function_decl(pair: Pair<Rule>) -> Option<Statement> {
    // Everything this function binds, so a LOCAL class declared in its body can
    // tell a captured value from one of its own members.
    ENCLOSING_LOCALS.with(|stack| stack.borrow_mut().push(bound_names(&pair)));
    let out = walk_function_decl_inner(pair);
    ENCLOSING_LOCALS.with(|stack| {
        stack.borrow_mut().pop();
    });
    out
}

fn walk_function_decl_inner(pair: Pair<Rule>) -> Option<Statement> {
    // An extension function reads the receiver's members unqualified, the same
    // as an extension property: `fun P.show() = "n=" + n`.
    let ext_receiver: Option<String> = pair
        .clone()
        .into_inner()
        .find(|p| p.as_rule() == Rule::receiver_prefix)
        .and_then(|p| p.into_inner().next())
        .map(|p| p.as_str().to_string());
    if let Some(receiver) = &ext_receiver {
        push_ext_receiver(receiver);
    }
    let out = walk_function_decl_body(pair);
    if ext_receiver.is_some() {
        pop_ext_receiver();
    }
    out
}

fn walk_function_decl_body(pair: Pair<Rule>) -> Option<Statement> {
    let mut name = String::new();
    let mut receiver_type: Option<String> = None;
    let mut params = Vec::new();
    let mut return_type = None;
    let mut body = Vec::new();

    let mut is_abstract = false;
    let mut is_operator = false;
    let mut visibility = Visibility::Public;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::modifier => {
                let m_str = inner.as_str();
                if m_str == "abstract" {
                    is_abstract = true;
                } else if m_str == "private" {
                    visibility = Visibility::Private;
                } else if m_str == "protected" {
                    visibility = Visibility::Protected;
                } else if m_str == "operator" {
                    is_operator = true;
                }
            }
            Rule::receiver_prefix => {
                if let Some(id_p) = inner.into_inner().next() {
                    receiver_type = Some(id_p.as_str().to_string());
                }
            }
            Rule::type_ref => {
                if return_type.is_none() && !name.is_empty() {
                    return_type = Some(type_hint_text(inner.as_str()));
                }
            }
            Rule::identifier => {
                name = inner.as_str().to_string();
            }
            Rule::parameter_list => {
                params = walk_parameter_list(inner);
            }
            Rule::function_body_expr => {
                if let Some(expr_pair) = inner.into_inner().next() {
                    let expr = walk_expr(expr_pair);
                    body.push(Statement::new(StmtKind::Return(Some(expr))));
                }
            }
            Rule::block => {
                body = walk_block_statements(inner);
            }
            _ => {}
        }
    }

    // `operator fun plus` is a DIFFERENT declaration from a plain `fun plus`:
    // only the former defines `+`. Kotlin's operator names are ordinary
    // identifiers, so the modifier is the only thing that distinguishes them
    // and it has to survive into `protocol.rs`, which decides slots. Encoded
    // in the name — the same device Dart uses for `operator+` — because the
    // slot mapping is a language-local decision and `Modifiers` is shared.
    // Stripped back off by `protocol::canonical_method`, so the member is
    // still stored under the name Kotlin code calls (`a.plus(b)` works).
    if is_operator && receiver_type.is_none() {
        name = format!("operator {}", name);
    }

    if receiver_type.is_some() {
        let mut ext_params = vec![Param {
            name: "this".to_string(),
            type_hint: receiver_type,
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        }];
        ext_params.extend(params);
        params = ext_params;
    }

    Some(Statement::new(StmtKind::FunctionDecl {
        name,
        params,
        return_type,
        body,
        modifiers: Modifiers {
            visibility,
            is_abstract,
            ..Default::default()
        },
        handles: vec![],
        is_async: false,
        is_generator: false,
        is_sub: false,
    }))
}

fn walk_parameter_list(pair: Pair<Rule>) -> Vec<Param> {
    let mut params = Vec::new();
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::parameter {
            let mut is_rest = false;
            let mut name = String::new();
            let mut type_hint = None;
            let mut default = None;
            let mut is_nullable = false;
            for p in inner.into_inner() {
                match p.as_rule() {
                    Rule::vararg_kw => is_rest = true,
                    Rule::identifier => name = p.as_str().to_string(),
                    Rule::type_ref => {
                        is_nullable = type_ref_is_nullable(p.as_str());
                        type_hint = Some(type_hint_text(p.as_str()));
                    }
                    Rule::expr => default = Some(walk_expr(p)),
                    _ => {}
                }
            }
            let is_optional = default.is_some();
            params.push(Param {
                name,
                type_hint,
                default,
                pass_by: PassBy::Value,
                is_rest,
                is_kwargs: false,
                is_optional,
                is_nullable,
            });
        }
    }
    params
}

fn walk_var_decl(pair: Pair<Rule>) -> Option<Statement> {
    if pair
        .clone()
        .into_inner()
        .any(|p| p.as_rule() == Rule::destructuring_target)
    {
        return walk_destructuring_decl(pair);
    }
    let mut is_readonly = false;
    let mut is_const = false;
    let mut name = String::new();
    let mut type_hint = None;
    let mut init = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::modifier => {
                if inner.as_str() == "const" {
                    is_const = true;
                }
            }
            Rule::val_kw => is_readonly = true,
            Rule::var_kw => is_readonly = false,
            Rule::identifier => name = inner.as_str().to_string(),
            Rule::type_ref => type_hint = Some(type_hint_text(inner.as_str())),
            Rule::expr => init = Some(walk_expr(inner)),
            _ => {}
        }
    }

    if type_hint.is_none() {
        if let Some(ref expr) = init {
            match expr.kind {
                ExprKind::Array(_) => type_hint = Some("Array".to_string()),
                ExprKind::Object(_) => type_hint = Some("Map".to_string()),
                _ => {}
            }
        }
    }

    Some(Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(name),
            type_hint,
            init,
            array_bounds: None,
            with_events: false,
        }],
        kind: if is_const || is_readonly {
            VarDeclKind::Const
        } else {
            VarDeclKind::Var
        },
    }))
}

/// Is this expression a STRING by construction, so `+` on it is
/// `kotlin.String.plus` (concatenation) rather than arithmetic?
///
/// Only syntactic evidence counts — a literal, a template, or a concatenation
/// already decided. Anything requiring the operand's runtime type is left to
/// the shared path, so this never claims a `+` it cannot prove.
fn kt_is_string_expr(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Lit(Literal::Str(_)) => true,
        ExprKind::Binary {
            op: BinOp::Concat, ..
        } => true,
        // `"$a$b".trimIndent()` and friends keep the template's type.
        ExprKind::Member { object, .. } => kt_is_string_expr(object),
        _ => false,
    }
}

/// A class-body `val`/`var` that declares `get()` / `set(v)` accessors.
///
/// Kotlin's properties are not fields: `val area: Int get() = w * h` has no
/// storage at all, and `var celsius` with a custom setter must run the setter
/// on assignment. The walker used to drop the accessors on the floor and emit a
/// plain `ClassMember::Field`, so the getter never ran and the property read as
/// `undefined`. `ClassMember::Property` is the model's own shape for this —
/// C# `{ get; set; }` and Pascal properties already use it, and the compiler
/// installs `__get_`/`__set_` accessors from it.
///
/// Returns `None` for a plain stored property, which stays a field.
thread_local! {
    /// The storage a `field` identifier means, while an accessor body is being
    /// walked. Empty everywhere else, so `field` stays an ordinary name.
    static BACKING_FIELD: std::cell::RefCell<Vec<String>> = const { std::cell::RefCell::new(Vec::new()) };
}

/// `field` inside an accessor is Kotlin's BACKING STORAGE, not a variable.
///
/// It is the only way a custom accessor can reach the property's own storage —
/// `set(v) { field = v + 1 }` — and reading it as a plain identifier left it
/// `undefined`. Every other identifier passes through untouched.
fn backing_field_substitution(name: &str) -> String {
    if name != "field" {
        return name.to_string();
    }
    BACKING_FIELD.with(|stack| {
        stack
            .borrow()
            .last()
            .cloned()
            .unwrap_or_else(|| name.to_string())
    })
}

/// The storage name a property's `field` resolves to. Distinct per property, so
/// two properties in one class each writing `field` do not share one slot.
fn backing_field_name(property: &str) -> String {
    format!("__kt_field_{}", property)
}

/// A `var_decl` that declares a PROPERTY, and the backing storage it needs.
///
/// Returns the `ClassMember::Property` plus, when the property has an
/// initializer or its accessors mention `field`, the private field that holds
/// its value. Kotlin's `var n: Int = 1  get() = field * 2` is NOT an
/// auto-property: the initializer seeds the BACKING FIELD and the accessors
/// still run. Reporting it as `auto_field` made the shared model treat it as a
/// plain field and drop both accessors (`classes.rs` skips auto-properties).
/// `val Receiver.name: T get() = …` — an extension PROPERTY.
///
/// It has no backing field: the accessor is the whole implementation, so it
/// lowers to a function of the receiver, exactly as an extension function does.
/// The read site (`x.name`) is rewritten to `name(x)`.
fn walk_extension_property(pair: Pair<Rule>) -> Option<Statement> {
    let inners: Vec<_> = pair.into_inner().collect();
    let receiver = inners
        .iter()
        .find(|p| p.as_rule() == Rule::receiver_prefix)?
        .clone()
        .into_inner()
        .next()?
        .as_str()
        .to_string();
    let name = inners
        .iter()
        .find(|p| p.as_rule() == Rule::identifier)?
        .as_str()
        .to_string();

    push_ext_receiver(&receiver);
    let mut body = Vec::new();
    for p in &inners {
        match p.as_rule() {
            Rule::property_accessor => {
                let mut is_get = false;
                for part in p.clone().into_inner() {
                    match part.as_rule() {
                        Rule::get_kw => is_get = true,
                        Rule::set_kw => is_get = false,
                        Rule::function_body_expr => {
                            if is_get {
                                if let Some(e) = part.into_inner().next() {
                                    body =
                                        vec![Statement::new(StmtKind::Return(Some(walk_expr(e))))];
                                }
                            }
                        }
                        Rule::block => {
                            if is_get {
                                body = walk_block_statements(part);
                            }
                        }
                        _ => {}
                    }
                }
            }
            Rule::expr if body.is_empty() => {
                body = vec![Statement::new(StmtKind::Return(Some(walk_expr(p.clone()))))];
            }
            _ => {}
        }
    }

    pop_ext_receiver();

    Some(Statement::new(StmtKind::FunctionDecl {
        name,
        params: vec![Param {
            name: "this".to_string(),
            type_hint: Some(receiver),
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        }],
        return_type: None,
        body,
        modifiers: Modifiers {
            visibility: Visibility::Public,
            is_extension: true,
            ..Default::default()
        },
        handles: vec![],
        is_async: false,
        is_generator: false,
        is_sub: false,
    }))
}

fn walk_class_property(pair: Pair<Rule>) -> Vec<ClassMember> {
    let inners: Vec<_> = pair.into_inner().collect();
    // `var x = 1` + `private set` declares an ORDINARY stored property whose
    // setter is restricted — the accessor has no body, so there is nothing to
    // run and the backing storage is the whole implementation. Only an accessor
    // WITH a body replaces the storage; treating a bodyless one as a computed
    // property left the field unreadable (`undefined`).
    let has_accessor_body = inners.iter().any(|p| {
        p.as_rule() == Rule::property_accessor
            && p.clone()
                .into_inner()
                .any(|part| matches!(part.as_rule(), Rule::function_body_expr | Rule::block))
    });
    if !has_accessor_body {
        return Vec::new();
    }

    // The property's own name, needed BEFORE the accessor bodies are walked so
    // `field` inside them resolves to this property's storage.
    let property_name = inners
        .iter()
        .find(|p| p.as_rule() == Rule::identifier)
        .map(|p| p.as_str().to_string())
        .unwrap_or_default();
    if property_name.is_empty() {
        return Vec::new();
    }
    let backing = backing_field_name(&property_name);
    let uses_field = inners.iter().any(|p| {
        p.as_rule() == Rule::property_accessor
            && p.as_str()
                .split(|c: char| !(c.is_ascii_alphanumeric() || c == '_'))
                .any(|word| word == "field")
    });
    BACKING_FIELD.with(|stack| stack.borrow_mut().push(backing.clone()));

    let mut name = String::new();
    let mut type_hint = None;
    let mut init = None;
    let mut is_readonly = false;
    let mut getter = None;
    let mut setter = None;

    for inner in inners {
        match inner.as_rule() {
            Rule::val_kw => is_readonly = true,
            Rule::identifier if name.is_empty() => name = inner.as_str().to_string(),
            Rule::type_ref => type_hint = Some(type_hint_text(inner.as_str())),
            Rule::expr => init = Some(walk_expr(inner)),
            Rule::property_accessor => {
                let mut is_get = false;
                let mut param_name = None;
                let mut body = Vec::new();
                for part in inner.into_inner() {
                    match part.as_rule() {
                        Rule::get_kw => is_get = true,
                        Rule::set_kw => is_get = false,
                        Rule::identifier => param_name = Some(part.as_str().to_string()),
                        // `get() = expr` is an expression body; the model wants
                        // statements, and the value of the accessor IS its
                        // result.
                        Rule::function_body_expr => {
                            if let Some(e) = part.into_inner().next() {
                                body = vec![Statement::new(StmtKind::Return(Some(walk_expr(e))))];
                            }
                        }
                        Rule::block => body = walk_block_statements(part),
                        _ => {}
                    }
                }
                if is_get {
                    getter = Some(body);
                } else {
                    setter = Some(PropertySetter {
                        param: Param {
                            // Kotlin's implicit setter parameter is `value`.
                            name: param_name.unwrap_or_else(|| "value".to_string()),
                            type_hint: None,
                            default: None,
                            pass_by: PassBy::Value,
                            is_rest: false,
                            is_kwargs: false,
                            is_optional: false,
                            is_nullable: false,
                        },
                        body,
                    });
                }
            }
            _ => {}
        }
    }

    BACKING_FIELD.with(|stack| {
        stack.borrow_mut().pop();
    });

    let mut out = Vec::new();
    // Storage exists when the source says it does: an initializer to hold, or
    // an accessor that reads or writes `field`.
    if init.is_some() || uses_field {
        out.push(ClassMember::Field {
            name: backing,
            type_hint: type_hint.clone(),
            init,
            modifiers: Modifiers {
                visibility: Visibility::Private,
                ..Default::default()
            },
            with_events: false,
            array_bounds: None,
        });
    }
    out.push(ClassMember::Property {
        name,
        type_hint,
        getter,
        setter,
        // Never an auto-property: these accessors have BODIES to run.
        is_auto: false,
        modifiers: Modifiers {
            visibility: Visibility::Public,
            is_readonly,
            ..Default::default()
        },
    });
    out
}

/// Whether `text` is a single Kotlin identifier and nothing else.
fn is_plain_identifier(text: &str) -> bool {
    let text = text.trim();
    !text.is_empty()
        && text
            .chars()
            .all(|c| c.is_ascii_alphanumeric() || c == '_' || c == '$')
        && !text.starts_with(|c: char| c.is_ascii_digit())
}

fn walk_class_decl(pair: Pair<Rule>) -> Option<Statement> {
    let mut name = String::new();
    let mut is_interface = false;
    let mut is_abstract = false;
    let mut is_sealed = false;

    let mut parents = Vec::new();
    let mut interfaces = Vec::new();
    let mut members = Vec::new();
    let mut base_args = None;
    let mut decorators = Vec::new();
    let mut init_stmts = Vec::new();

    // A LOCAL class reads values from the function around it. With
    // `implicit_self_fields` a bare identifier in a method is a FIELD read, so
    // those reads never reached the enclosing scope and answered `undefined`.
    // Kotlin gives a local class synthetic storage for what it captures; that
    // lowering is declared here, as leading constructor parameters.
    let captures: Vec<String> = ENCLOSING_LOCALS.with(|stack| {
        let stack = stack.borrow();
        let Some(enclosing) = stack.last() else {
            return Vec::new();
        };
        let declared = bound_names(&pair);
        let mut own: std::collections::HashSet<String> = declared;
        for sub in pair.clone().into_inner() {
            if sub.as_rule() == Rule::class_body {
                for m in sub.into_inner() {
                    if let Some(inner) = m.into_inner().next() {
                        if inner.as_rule() == Rule::function_decl {
                            if let Some(id) = inner
                                .clone()
                                .into_inner()
                                .find(|p| p.as_rule() == Rule::identifier)
                            {
                                own.insert(id.as_str().to_string());
                            }
                        }
                    }
                }
            }
        }
        let mut caps: Vec<String> = read_names(&pair)
            .into_iter()
            .filter(|n| enclosing.contains(n) && !own.contains(n))
            .collect();
        caps.sort();
        caps
    });

    let inner_pairs: Vec<_> = pair.into_inner().collect();
    if let Some(id) = inner_pairs.iter().find(|p| p.as_rule() == Rule::identifier) {
        name = id.as_str().to_string();
    }
    let is_inner_class = inner_pairs
        .iter()
        .any(|p| p.as_rule() == Rule::modifier && p.as_str().trim() == "inner");
    let enclosing_classes = CURRENT_CLASS_STACK.with(|stack| stack.borrow().clone());
    let stack_class_name = enclosing_classes
        .last()
        .map(|outer| format!("{outer}.{name}"))
        .unwrap_or_else(|| name.clone());
    let outer_context = if is_inner_class {
        CLASS_MEMBERS.with(|members| {
            let members = members.borrow();
            enclosing_classes
                .iter()
                .rev()
                .map(|class_name| {
                    (
                        class_name.clone(),
                        members.get(class_name).cloned().unwrap_or_default(),
                    )
                })
                .collect::<Vec<_>>()
        })
    } else {
        Vec::new()
    };
    if !outer_context.is_empty() {
        INNER_OUTER_MEMBERS.with(|stack| stack.borrow_mut().push(outer_context));
    }
    if !name.is_empty() {
        CURRENT_CLASS_STACK.with(|stack| stack.borrow_mut().push(stack_class_name));
    }

    let mut is_data = false;
    let mut primary_prop_names = Vec::new();
    // Which member is the PRIMARY constructor, so property initializers and
    // `init` blocks land AFTER its parameter-to-field assignments.
    let mut primary_ctor_index: Option<usize> = None;
    // `: I by <expr>` — the field the forwarders read, and what fills it.
    // Kotlin stores the delegate; `AugmentDecl::via_field` names STORAGE, so a
    // `by` whose expression is not already a property needs one declared.
    let mut delegate_storage: Vec<(String, Expression)> = Vec::new();

    for inner in &inner_pairs {
        if inner.as_rule() == Rule::inheritance_list {
            for spec in inner.clone().into_inner() {
                if spec.as_rule() == Rule::inheritance_specifier {
                    let mut parent_name = String::new();
                    let mut spec_base_args = Vec::new();
                    let mut by_expr = None;
                    // Kotlin marks the SUPERCLASS by calling its constructor:
                    // `class D : B(n), I` extends `B` and implements `I`. An
                    // interface is never constructed, so parentheses are the
                    // whole distinction — and `B()` with no arguments has to
                    // count too, which an empty `arg_list` cannot express.
                    // Taking the FIRST supertype as the parent instead made
                    // `class C : A, B` extend `A`, so `B`'s members vanished.
                    let calls_constructor = spec.as_str().contains('(');
                    for sub in spec.into_inner() {
                        match sub.as_rule() {
                            // `type_ref` is non-atomic, so its span carries any
                            // trailing whitespace before `by` / `(`. The name is
                            // a LOOKUP KEY — `available.get("Greeter ")` misses
                            // every time — so it has to be trimmed here.
                            Rule::type_ref => parent_name = sub.as_str().trim().to_string(),
                            Rule::arg_list => {
                                for arg_p in sub.into_inner() {
                                    let mut arg_expr = None;
                                    for e in arg_p.into_inner() {
                                        if e.as_rule() == Rule::expr {
                                            arg_expr = Some(walk_expr(e));
                                        }
                                    }
                                    if let Some(ae) = arg_expr {
                                        spec_base_args.push(ae);
                                    }
                                }
                            }
                            Rule::delegate_expr => {
                                by_expr = Some((sub.as_str().to_string(), walk_expr(sub)))
                            }
                            _ => {}
                        }
                    }
                    if !parent_name.is_empty() {
                        if let Some((text, expr)) = by_expr {
                            // The delegate lives in a FIELD. A bare identifier
                            // names one directly (`by base`); anything else
                            // (`by ArrayList()`, `by makeIt()`) is an
                            // expression Kotlin evaluates once into a synthetic
                            // property, so declare that property here — a
                            // forwarder reading `this.ArrayList()` is not a
                            // read at all.
                            let field = if is_plain_identifier(&text) {
                                text.trim().to_string()
                            } else {
                                format!("__kt_delegate_{}", parent_name.replace('.', "_"))
                            };
                            delegate_storage.push((field.clone(), expr));
                            // A delegating class IS the interface — `is I` and
                            // every interface-typed binding depend on it.
                            interfaces.push(parent_name.clone());
                            members.push(ClassMember::Augment(AugmentDecl {
                                from: parent_name.clone(),
                                via_field: Some(field),
                                adjustments: vec![],
                            }));
                        } else if calls_constructor && parents.is_empty() && !is_interface {
                            parents.push(parent_name);
                            if !spec_base_args.is_empty() {
                                base_args = Some(spec_base_args);
                            }
                        } else {
                            // An implemented interface contributes its DEFAULT
                            // methods. `interfaces` alone cannot do that — the
                            // model's own doc says that list is for identity
                            // checks and "method dispatch never walks it" — so
                            // the contribution is declared as an augmentation
                            // and the shared `class_augmentation` pass applies
                            // it, which is flexclassplan §4c's "Java interface
                            // default methods" arriving as one model rather
                            // than a fifth walker fold.
                            interfaces.push(parent_name.clone());
                            members.push(ClassMember::Augment(AugmentDecl {
                                from: parent_name,
                                via_field: None,
                                adjustments: vec![],
                            }));
                        }
                    }
                }
            }
        }
    }

    // The body walk needs to know which supertype is the SUPERCLASS: a
    // `super<X>.f()` whose `X` is the parent is an ordinary super call, and one
    // whose `X` is an implemented interface is a default reached by alias.
    CURRENT_CLASS_PARENT.with(|stack| stack.borrow_mut().push(parents.first().cloned()));
    SUPER_QUALIFIED_USES.with(|stack| stack.borrow_mut().push(Vec::new()));

    for inner in inner_pairs {
        match inner.as_rule() {
            Rule::annotation => decorators.push(walk_annotation(inner)),
            Rule::interface_kw => is_interface = true,
            Rule::modifier => match inner.as_str() {
                "abstract" => is_abstract = true,
                "sealed" => is_sealed = true,
                "data" => is_data = true,
                _ => {}
            },
            Rule::identifier => name = inner.as_str().to_string(),
            Rule::primary_constructor => {
                let mut ctor_params = Vec::new();
                let mut ctor_body = Vec::new();

                for param in inner.into_inner() {
                    if param.as_rule() == Rule::class_parameter {
                        let mut param_is_prop = false;
                        let mut is_readonly = false;
                        let mut pname = String::new();
                        let mut type_hint = None;
                        // `class Point(val x: Int = 0)` — the default is part of
                        // the primary constructor's signature, and `copy()`
                        // re-states it. Dropping it made every call that omitted
                        // the argument bind `undefined`.
                        let mut default = None;
                        for p in param.into_inner() {
                            match p.as_rule() {
                                Rule::val_kw => {
                                    param_is_prop = true;
                                    is_readonly = true;
                                }
                                Rule::var_kw => {
                                    param_is_prop = true;
                                    is_readonly = false;
                                }
                                Rule::identifier => pname = p.as_str().to_string(),
                                Rule::type_ref => type_hint = Some(type_hint_text(p.as_str())),
                                Rule::expr => default = Some(walk_expr(p.clone())),
                                _ => {}
                            }
                        }
                        if !pname.is_empty() {
                            primary_prop_names.push(pname.clone());
                            let is_optional = default.is_some();
                            ctor_params.push(Param {
                                name: pname.clone(),
                                type_hint: type_hint.clone(),
                                default,
                                pass_by: PassBy::Value,
                                is_rest: false,
                                is_kwargs: false,
                                is_optional,
                                is_nullable: false,
                            });
                            if param_is_prop {
                                members.push(ClassMember::Field {
                                    name: pname.clone(),
                                    type_hint: type_hint.clone(),
                                    init: None,
                                    modifiers: Modifiers {
                                        visibility: Visibility::Public,
                                        is_readonly,
                                        ..Default::default()
                                    },
                                    with_events: false,
                                    array_bounds: None,
                                });
                                ctor_body.push(Statement::new(StmtKind::Expr(Expression::new(
                                    ExprKind::Assign {
                                        target: Box::new(Expression::new(ExprKind::Member {
                                            object: Box::new(Expression::new(ExprKind::This)),
                                            field: pname.clone(),
                                            null_safe: false,
                                        })),
                                        value: Box::new(Expression::ident(&pname)),
                                    },
                                ))));
                                let prop_idx = (primary_prop_names.len() - 1) as i64;
                                ctor_body.push(Statement::new(StmtKind::Expr(Expression::new(
                                    ExprKind::Assign {
                                        target: Box::new(Expression::new(ExprKind::Index {
                                            object: Box::new(Expression::new(ExprKind::This)),
                                            index: Box::new(Expression::int(prop_idx)),
                                            null_safe: false,
                                        })),
                                        value: Box::new(Expression::ident(&pname)),
                                    },
                                ))));
                            }
                        }
                    }
                }

                primary_ctor_index = Some(members.len());
                members.push(ClassMember::Constructor {
                    name: None,
                    params: ctor_params,
                    body: ctor_body,
                    base_args: base_args.clone(),
                    initializer_target: ConstructorInitializerTarget::Base,
                    visibility: Visibility::Public,
                });
            }
            Rule::class_body => {
                for member_pair in inner.into_inner() {
                    if member_pair.as_rule() == Rule::class_member {
                        if let Some(inner_member) = member_pair.into_inner().next() {
                            match inner_member.as_rule() {
                                Rule::init_block => {
                                    if let Some(block_pair) = inner_member.into_inner().next() {
                                        init_stmts.extend(walk_block_statements(block_pair));
                                    }
                                }
                                Rule::secondary_constructor => {
                                    let mut s_params = Vec::new();
                                    let mut s_body = Vec::new();
                                    let mut s_target = ConstructorInitializerTarget::Base;
                                    let mut s_base_args = None;
                                    for sc in inner_member.into_inner() {
                                        match sc.as_rule() {
                                            Rule::parameter_list => {
                                                s_params = walk_parameter_list(sc)
                                            }
                                            Rule::this_kw => {
                                                s_target = ConstructorInitializerTarget::This
                                            }
                                            Rule::super_kw => {
                                                s_target = ConstructorInitializerTarget::Base
                                            }
                                            Rule::arg_list => {
                                                let mut bargs = Vec::new();
                                                for arg_p in sc.into_inner() {
                                                    for e in arg_p.into_inner() {
                                                        if e.as_rule() == Rule::expr {
                                                            bargs.push(walk_expr(e));
                                                        }
                                                    }
                                                }
                                                s_base_args = Some(bargs);
                                            }
                                            Rule::block => s_body = walk_block_statements(sc),
                                            _ => {}
                                        }
                                    }
                                    members.push(ClassMember::Constructor {
                                        name: None,
                                        params: s_params,
                                        body: s_body,
                                        base_args: s_base_args,
                                        initializer_target: s_target,
                                        visibility: Visibility::Public,
                                    });
                                }
                                Rule::class_decl | Rule::object_decl | Rule::interface_decl => {
                                    if let Some(stmt) = walk_statement(inner_member) {
                                        members.push(ClassMember::NestedType(Box::new(stmt)));
                                    }
                                }
                                Rule::function_decl => {
                                    if let Some(stmt) = walk_function_decl(inner_member) {
                                        members.push(ClassMember::Method(Box::new(stmt)));
                                    }
                                }
                                Rule::var_decl => {
                                    // `val area: Int get() = w * h` is a
                                    // PROPERTY, not a field — it has no storage
                                    // and the accessor has to run on each read.
                                    let prop = walk_class_property(inner_member.clone());
                                    if !prop.is_empty() {
                                        members.extend(prop);
                                        continue;
                                    }
                                    // `val` is READ-ONLY, not `const`. Only
                                    // `const val` is a compile-time constant
                                    // with static storage; a plain `val n = 5`
                                    // is an instance property, and routing it
                                    // to `ClassMember::Const` gave EVERY `val`
                                    // property in a class body static storage —
                                    // `class B { val zz = "b" }` then read
                                    // `B().zz` as `undefined`. `VarDeclKind`
                                    // collapses the two, so ask the source.
                                    let is_const_val = inner_member.clone().into_inner().any(|p| {
                                        p.as_rule() == Rule::modifier
                                            && p.as_str().trim() == "const"
                                    });
                                    let is_readonly_val = inner_member
                                        .clone()
                                        .into_inner()
                                        .any(|p| p.as_rule() == Rule::val_kw);
                                    if let Some(stmt) = walk_var_decl(inner_member) {
                                        if let StmtKind::VarDecl { declarations, .. } = stmt.kind {
                                            for decl in declarations {
                                                if let BindingPattern::Ident(fname) = decl.pattern {
                                                    if is_const_val {
                                                        if let Some(val_expr) = decl.init {
                                                            members.push(ClassMember::Const {
                                                                name: fname,
                                                                type_hint: decl.type_hint,
                                                                value: val_expr,
                                                                visibility: Visibility::Public,
                                                            });
                                                        }
                                                    } else {
                                                        members.push(ClassMember::Field {
                                                            name: fname,
                                                            type_hint: decl.type_hint,
                                                            init: decl.init,
                                                            modifiers: Modifiers {
                                                                visibility: Visibility::Public,
                                                                is_readonly: is_readonly_val,
                                                                ..Default::default()
                                                            },
                                                            with_events: false,
                                                            array_bounds: None,
                                                        });
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                Rule::companion_object => {
                                    if let Some(stmt) = walk_object_decl(inner_member) {
                                        if let StmtKind::ClassDecl {
                                            members: comp_members,
                                            ..
                                        } = stmt.kind
                                        {
                                            for mut cm in comp_members {
                                                if let ClassMember::Method(ref mut mstmt) = cm {
                                                    if let StmtKind::FunctionDecl {
                                                        ref mut modifiers,
                                                        ..
                                                    } = mstmt.kind
                                                    {
                                                        modifiers.is_static = true;
                                                    }
                                                }
                                                members.push(cm);
                                            }
                                        }
                                    }
                                }
                                _ => {}
                            }
                        }
                    }
                }
            }
            _ => {}
        }
    }

    // A `data class` synthesizes NOTHING here. Its members are DERIVED from
    // the primary constructor, which is `normalize_class.rs`'s job — the
    // walker only DECLARES what the source said (flexclassplan §0.3). This
    // block used to hand-build `componentN`, `copy`, `toString`, `equals` and
    // `hashCode` as raw AST, a second class implementation sitting beside the
    // one `classes.rs` already has (§0.1, §4a-ter).
    CURRENT_CLASS_PARENT.with(|stack| {
        stack.borrow_mut().pop();
    });
    if !name.is_empty() {
        CURRENT_CLASS_STACK.with(|stack| {
            stack.borrow_mut().pop();
        });
    }
    if is_inner_class && !enclosing_classes.is_empty() {
        INNER_OUTER_MEMBERS.with(|stack| {
            stack.borrow_mut().pop();
        });
    }
    // Every `super<I>.m()` the body used becomes an ADDITIVE alias on `I`'s
    // augmentation. `bound_names` in the shared fold binds a renamed member
    // under both its own name and the alias, so the class keeps its own `m`
    // and still reaches the interface default.
    let super_uses = SUPER_QUALIFIED_USES
        .with(|stack| stack.borrow_mut().pop())
        .unwrap_or_default();
    for (from, member) in super_uses {
        for m in &mut members {
            if let ClassMember::Augment(decl) = m {
                if decl.from == from
                    && !decl
                        .adjustments
                        .iter()
                        .any(|a| a.member == member && a.rename_to.is_some())
                {
                    decl.adjustments.push(AugmentAdjustment {
                        member: member.clone(),
                        rename_to: Some(super_alias(&from, &member)),
                        ..Default::default()
                    });
                }
            }
        }
    }

    // Declare the captures: a private field each, a leading constructor
    // parameter each, and the assignment that binds them. The construction site
    // passes them (see `LOCAL_CLASS_CAPTURES` at the `New` site).
    if !captures.is_empty() && !name.is_empty() {
        LOCAL_CLASS_CAPTURES.with(|m| m.borrow_mut().insert(name.clone(), captures.clone()));
        for cap in &captures {
            members.push(ClassMember::Field {
                name: cap.clone(),
                type_hint: None,
                init: None,
                modifiers: Modifiers {
                    visibility: Visibility::Private,
                    is_readonly: true,
                    ..Default::default()
                },
                with_events: false,
                array_bounds: None,
            });
        }
        let capture_params: Vec<Param> = captures
            .iter()
            .map(|cap| Param {
                name: cap.clone(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: true,
            })
            .collect();
        let capture_assigns: Vec<Statement> = captures
            .iter()
            .map(|cap| {
                Statement::new(StmtKind::Expr(Expression::new(ExprKind::Assign {
                    target: Box::new(Expression::new(ExprKind::Member {
                        object: Box::new(Expression::new(ExprKind::This)),
                        field: cap.clone(),
                        null_safe: false,
                    })),
                    value: Box::new(Expression::ident(cap)),
                })))
            })
            .collect();
        if !members
            .iter()
            .any(|m| matches!(m, ClassMember::Constructor { .. }))
        {
            primary_ctor_index = Some(members.len());
            members.push(ClassMember::Constructor {
                name: None,
                params: Vec::new(),
                body: Vec::new(),
                base_args: base_args.clone(),
                initializer_target: ConstructorInitializerTarget::Base,
                visibility: Visibility::Public,
            });
        }
        for member in &mut members {
            if let ClassMember::Constructor { params, body, .. } = member {
                let mut p = capture_params.clone();
                p.extend(std::mem::take(params));
                *params = p;
                let mut b = capture_assigns.clone();
                b.extend(std::mem::take(body));
                *body = b;
            }
        }
    }

    if is_inner_class && !enclosing_classes.is_empty() {
        let outer_name = enclosing_classes.last().cloned();
        members.push(ClassMember::Field {
            name: "__kt_outer".to_string(),
            type_hint: outer_name.clone(),
            init: Some(Expression::ident("__kt_outer")),
            modifiers: Modifiers {
                visibility: Visibility::Private,
                is_readonly: true,
                ..Default::default()
            },
            with_events: false,
            array_bounds: None,
        });
        let outer_param = Param {
            name: "__kt_outer".to_string(),
            type_hint: outer_name,
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        };
        if !members
            .iter()
            .any(|m| matches!(m, ClassMember::Constructor { .. }))
        {
            primary_ctor_index = Some(members.len());
            members.push(ClassMember::Constructor {
                name: None,
                params: Vec::new(),
                body: Vec::new(),
                base_args: base_args.clone(),
                initializer_target: ConstructorInitializerTarget::Base,
                visibility: Visibility::Public,
            });
        }
        for member in &mut members {
            if let ClassMember::Constructor { params, body, .. } = member {
                params.insert(0, outer_param.clone());
            }
        }
    }

    // `class C : P { constructor() : super(1) { … } }` — no parentheses on `P`,
    // because a SECONDARY constructor calls the base constructor instead. The
    // parenthesis rule alone read `P` as an interface, so `C` had no superclass
    // and every inherited member came back `undefined`. A `super(…)` delegation
    // is the same statement of intent that `: P(…)` makes.
    let delegates_to_super = members.iter().any(|m| {
        matches!(
            m,
            ClassMember::Constructor {
                initializer_target: ConstructorInitializerTarget::Base,
                base_args: Some(_),
                ..
            }
        )
    });
    if delegates_to_super && parents.is_empty() && !interfaces.is_empty() {
        // Kotlin requires the superclass to be listed first.
        let base = interfaces.remove(0);
        members.retain(
            |m| !matches!(m, ClassMember::Augment(d) if d.from == base && d.via_field.is_none()),
        );
        parents.push(base);
    }

    // Give every `by` delegate its storage. A `val`/`var` primary-constructor
    // parameter already declared the field; a PLAIN parameter (`class B(base:
    // Counter) : Counter by base`) did not, and the forwarders read
    // `this.base` — which is why a delegate that was not also a property came
    // out `undefined is not callable`.
    for (field, init) in &delegate_storage {
        if members
            .iter()
            .any(|m| matches!(m, ClassMember::Field { name, .. } if name == field))
        {
            continue;
        }
        let from_ctor_param = members.iter().any(|m| {
            matches!(m, ClassMember::Constructor { params, .. }
                if params.iter().any(|p| &p.name == field))
        });
        members.push(ClassMember::Field {
            name: field.clone(),
            type_hint: None,
            // A constructor parameter is only in scope INSIDE the constructor,
            // so it is assigned there; anything else is an expression the class
            // owns outright and initialises with the field.
            init: if from_ctor_param {
                None
            } else {
                Some(init.clone())
            },
            modifiers: Modifiers {
                visibility: Visibility::Private,
                is_readonly: true,
                ..Default::default()
            },
            with_events: false,
            array_bounds: None,
        });
        if from_ctor_param {
            init_stmts.push(Statement::new(StmtKind::Expr(Expression::new(
                ExprKind::Assign {
                    target: Box::new(Expression::new(ExprKind::Member {
                        object: Box::new(Expression::new(ExprKind::This)),
                        field: field.clone(),
                        null_safe: false,
                    })),
                    value: Box::new(Expression::ident(field)),
                },
            ))));
        }
    }

    if !init_stmts.is_empty() {
        // Property initializers and `init` blocks run in the CONSTRUCTOR. A
        // class with neither a primary constructor nor a secondary one has no
        // `ClassMember::Constructor` for them to merge into, and they were
        // silently dropped — `class C { val n = 5 }` left `n` undefined. Give
        // that class the no-arg constructor Kotlin gives it.
        if !members
            .iter()
            .any(|m| matches!(m, ClassMember::Constructor { .. }))
        {
            members.push(ClassMember::Constructor {
                name: None,
                params: Vec::new(),
                body: Vec::new(),
                base_args: base_args.clone(),
                initializer_target: ConstructorInitializerTarget::Base,
                visibility: Visibility::Public,
            });
        }
        for (idx, member) in members.iter_mut().enumerate() {
            if let ClassMember::Constructor {
                body,
                initializer_target,
                ..
            } = member
            {
                // A secondary constructor that delegates with `: this(…)` must
                // NOT re-run them — the constructor it delegates to already
                // did, and running them twice re-initialises every property
                // after the delegate set it.
                if *initializer_target == ConstructorInitializerTarget::This {
                    continue;
                }
                // Kotlin runs the primary constructor's parameter bindings
                // FIRST, then property initializers and `init` blocks in source
                // order. Putting the initializers first made `class K(val x:
                // Int) { init { println(this.x) } }` print `undefined` — the
                // init block ran before `x` was ever assigned. A SECONDARY
                // constructor is the other way round: the initializers are part
                // of the object's construction and its own body runs last.
                let mut combined = if Some(idx) == primary_ctor_index {
                    std::mem::take(body)
                } else {
                    Vec::new()
                };
                combined.extend(init_stmts.clone());
                if Some(idx) != primary_ctor_index {
                    combined.extend(std::mem::take(body));
                }
                *body = combined;
            }
        }
    }

    Some(Statement::new(StmtKind::ClassDecl {
        name,
        parents,
        interfaces,
        members,
        modifiers: ClassModifiers {
            is_abstract: is_abstract || is_interface,
            is_sealed,
            // What this declaration DECLARES. `emit_class_from_ast` copies it
            // into `NormalClass.declared_kind` and the compiler stamps it on
            // the class object — which is what answers `interface_exists` and
            // keeps an interface from being treated as an instantiable class.
            // Left at the `Class` default, every Kotlin `interface` claimed to
            // be one.
            kind: if is_interface {
                ClassKind::Interface
            } else if is_data {
                // `data` is a DECLARED KIND, not a walker instruction. Stating
                // it here is the whole of the frontend's job; normalization
                // derives the members.
                ClassKind::Record
            } else {
                ClassKind::Class
            },
            ..Default::default()
        },
        decorators,
    }))
}

/// The members of a `class_body` that belongs to an OBJECT — a named `object`,
/// a `companion object`, or an anonymous `object : I { … }`.
///
/// One walk for all three. Each of them used to walk only `function_decl`, so
/// every `val`/`var` an object declared was dropped: `companion object { val K
/// = 7 }` read `C.K` as `undefined` and an anonymous object's state did not
/// exist at all.
///
/// `statics` — a named object and a companion are SINGLETONS, so their storage
/// is static; an anonymous object is an ordinary instance.
fn walk_object_body_members(class_body: Pair<Rule>, statics: bool) -> Vec<ClassMember> {
    let mut members = Vec::new();
    for member_pair in class_body.into_inner() {
        if member_pair.as_rule() != Rule::class_member {
            continue;
        }
        let Some(inner_member) = member_pair.into_inner().next() else {
            continue;
        };
        match inner_member.as_rule() {
            Rule::function_decl => {
                if let Some(mut stmt) = walk_function_decl(inner_member) {
                    if let StmtKind::FunctionDecl {
                        ref mut modifiers, ..
                    } = stmt.kind
                    {
                        modifiers.is_static = statics;
                    }
                    members.push(ClassMember::Method(Box::new(stmt)));
                }
            }
            Rule::var_decl => {
                if !statics {
                    // An accessor-backed property has no storage; the shared
                    // `ClassMember::Property` carries it.
                    let prop = walk_class_property(inner_member.clone());
                    if !prop.is_empty() {
                        members.extend(prop);
                        continue;
                    }
                }
                let is_const_val = inner_member
                    .clone()
                    .into_inner()
                    .any(|p| p.as_rule() == Rule::modifier && p.as_str().trim() == "const");
                let is_readonly_val = inner_member
                    .clone()
                    .into_inner()
                    .any(|p| p.as_rule() == Rule::val_kw);
                let Some(stmt) = walk_var_decl(inner_member) else {
                    continue;
                };
                let StmtKind::VarDecl { declarations, .. } = stmt.kind else {
                    continue;
                };
                for decl in declarations {
                    let BindingPattern::Ident(fname) = decl.pattern else {
                        continue;
                    };
                    if is_const_val {
                        if let Some(value) = decl.init {
                            members.push(ClassMember::Const {
                                name: fname,
                                type_hint: decl.type_hint,
                                value,
                                visibility: Visibility::Public,
                            });
                        }
                        continue;
                    }
                    members.push(ClassMember::Field {
                        name: fname,
                        type_hint: decl.type_hint,
                        init: decl.init,
                        modifiers: Modifiers {
                            visibility: Visibility::Public,
                            is_static: statics,
                            is_readonly: is_readonly_val,
                            ..Default::default()
                        },
                        with_events: false,
                        array_bounds: None,
                    });
                }
            }
            Rule::class_decl | Rule::object_decl | Rule::interface_decl => {
                if let Some(stmt) = walk_statement(inner_member) {
                    members.push(ClassMember::NestedType(Box::new(stmt)));
                }
            }
            _ => {}
        }
    }
    members
}

fn walk_object_decl(pair: Pair<Rule>) -> Option<Statement> {
    let mut name = "Companion".to_string();
    let mut members = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::identifier => name = inner.as_str().to_string(),
            Rule::class_body => {
                members.extend(walk_object_body_members(inner, true));
            }
            _ => {}
        }
    }

    Some(Statement::new(StmtKind::ClassDecl {
        name,
        parents: vec![],
        interfaces: vec![],
        members,
        modifiers: ClassModifiers::default(),
        decorators: vec![],
    }))
}

fn walk_block_statements(pair: Pair<Rule>) -> Vec<Statement> {
    let mut stmts = Vec::new();
    for inner in pair.into_inner() {
        if let Some(stmt) = walk_statement(inner) {
            stmts.push(stmt);
        }
    }
    stmts
}

fn walk_if_stmt(pair: Pair<Rule>) -> Option<Statement> {
    let mut cond = Expression::null();
    let mut then_body = Vec::new();
    let mut else_body = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::expr => cond = walk_expr(p),
            Rule::block => {
                if then_body.is_empty() {
                    then_body = walk_block_statements(p);
                } else {
                    else_body = Some(walk_block_statements(p));
                }
            }
            Rule::statement => {
                if then_body.is_empty() {
                    if let Some(s) = walk_statement(p) {
                        then_body = vec![s];
                    }
                } else {
                    if let Some(s) = walk_statement(p) {
                        else_body = Some(vec![s]);
                    }
                }
            }
            Rule::if_expr => {
                if let Some(s) = walk_if_stmt(p) {
                    else_body = Some(vec![s]);
                }
            }
            _ => {}
        }
    }

    Some(Statement::new(StmtKind::If {
        cond,
        then_body,
        elifs: vec![],
        else_body,
    }))
}

/// Whether this `when` condition is a PREDICATE rather than a value to compare
/// the subject against.
///
/// `is T`, `!is T`, `in c` and `!in c` test the subject; they are not values it
/// could equal. `CaseCondition` (statement `when`) and `MatchArm::conditions`
/// (expression `when`) both hold values only, so a `when` containing any of
/// these has to be emitted in the other shape Kotlin already has — see
/// [`when_condition_predicate`].
fn when_condition_is_predicate(pair: &Pair<Rule>) -> bool {
    pair.clone().into_inner().any(|sub| {
        matches!(
            sub.as_rule(),
            Rule::is_kw | Rule::not_is_kw | Rule::in_kw | Rule::not_in_kw
        )
    })
}

/// As [`when_condition_is_predicate`], plus the two forms only the STATEMENT
/// shape can carry: `MatchArm` has no range and no bare comparison, so an
/// expression `when` using either has to go predicate-shaped as well.
fn when_condition_needs_predicate_expr(pair: &Pair<Rule>) -> bool {
    when_condition_is_predicate(pair)
        || pair.clone().into_inner().any(|sub| {
            matches!(
                sub.as_rule(),
                Rule::range_condition | Rule::comparison_condition
            )
        })
}

/// One `when` condition rendered as a BOOLEAN test of `subject`.
///
/// This is the subjectless shape — `when { cond -> … }`, discriminant `true`,
/// every condition a bool — which Kotlin has and the compiler already lowers.
/// Reusing it is what keeps `is`/`in`/range/comparison arms out of a second
/// switch lowering of their own.
fn when_condition_predicate(pair: Pair<Rule>, subject: &Expression) -> Option<Expression> {
    let mut prefix: Option<Rule> = None;
    let mut out: Option<Expression> = None;

    for sub in pair.into_inner() {
        match sub.as_rule() {
            Rule::is_kw | Rule::not_is_kw | Rule::in_kw | Rule::not_in_kw => {
                prefix = Some(sub.as_rule())
            }
            Rule::type_ref => {
                // Lowercased to match the `is` operator's own walk above —
                // `IsType` names are compared case-folded.
                let test = Expression::new(ExprKind::IsType {
                    expr: Box::new(subject.clone()),
                    type_name: type_hint_text(sub.as_str()).to_lowercase(),
                });
                out = Some(if prefix == Some(Rule::not_is_kw) {
                    not_expr(test)
                } else {
                    test
                });
            }
            Rule::range_condition => {
                let mut bounds = sub.into_inner();
                if let (Some(lo), Some(hi)) = (bounds.next(), bounds.next()) {
                    let lower = Expression::new(ExprKind::Binary {
                        op: BinOp::GtEq,
                        left: Box::new(subject.clone()),
                        right: Box::new(walk_expr(lo)),
                    });
                    let upper = Expression::new(ExprKind::Binary {
                        op: BinOp::LtEq,
                        left: Box::new(subject.clone()),
                        right: Box::new(walk_expr(hi)),
                    });
                    out = Some(Expression::new(ExprKind::Binary {
                        op: BinOp::And,
                        left: Box::new(lower),
                        right: Box::new(upper),
                    }));
                }
            }
            Rule::comparison_condition => {
                let op_str = sub.as_str();
                let op = if op_str.starts_with(">=") {
                    BinOp::GtEq
                } else if op_str.starts_with("<=") {
                    BinOp::LtEq
                } else if op_str.starts_with('>') {
                    BinOp::Gt
                } else if op_str.starts_with('<') {
                    BinOp::Lt
                } else if op_str.starts_with("!=") {
                    BinOp::NotEq
                } else {
                    BinOp::Eq
                };
                if let Some(rhs) = sub.into_inner().next() {
                    out = Some(Expression::new(ExprKind::Binary {
                        op,
                        left: Box::new(subject.clone()),
                        right: Box::new(walk_expr(rhs)),
                    }));
                }
            }
            Rule::expr => {
                let value = walk_expr(sub);
                let membership = |negated: bool| {
                    let test = Expression::new(ExprKind::Binary {
                        op: BinOp::In,
                        left: Box::new(subject.clone()),
                        right: Box::new(value.clone()),
                    });
                    if negated { not_expr(test) } else { test }
                };
                out = Some(match prefix {
                    Some(Rule::in_kw) => membership(false),
                    Some(Rule::not_in_kw) => membership(true),
                    _ => Expression::new(ExprKind::Binary {
                        op: BinOp::Eq,
                        left: Box::new(subject.clone()),
                        right: Box::new(value),
                    }),
                });
            }
            _ => {}
        }
    }

    out
}

fn not_expr(expr: Expression) -> Expression {
    Expression::new(ExprKind::Unary {
        op: UnaryOp::Not,
        expr: Box::new(expr),
    })
}

fn walk_when_stmt(pair: Pair<Rule>) -> Option<Statement> {
    let mut disc = None;
    let mut entries = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::expr => disc = Some(walk_expr(p)),
            Rule::when_entry => entries.push(p),
            _ => {}
        }
    }

    // A `when` whose subject is TESTED (`is`, `in`) rather than compared can't
    // be carried by `CaseCondition`, so the whole statement flips to the
    // subjectless shape: discriminant `true`, every condition a boolean test of
    // the subject. All or nothing — one switch has one discriminant.
    let predicate_mode = disc.is_some()
        && entries.iter().any(|entry| {
            entry
                .clone()
                .into_inner()
                .any(|p| p.as_rule() == Rule::when_condition && when_condition_is_predicate(&p))
        });
    let subject = disc.clone().unwrap_or_else(|| Expression::bool(true));

    let mut cases = Vec::new();
    let mut default = None;

    for entry in entries {
        let mut entry_inner = entry.into_inner();
        let mut is_else = false;
        let mut cond_cases = Vec::new();
        let mut body_stmts = Vec::new();

        while let Some(p) = entry_inner.next() {
            match p.as_rule() {
                Rule::else_kw => is_else = true,
                Rule::when_condition if predicate_mode => {
                    if let Some(test) = when_condition_predicate(p, &subject) {
                        cond_cases.push(CaseCondition::Value(test));
                    }
                }
                Rule::when_condition => {
                    for csub in p.into_inner() {
                        match csub.as_rule() {
                            Rule::range_condition => {
                                let mut r_exprs = csub.into_inner();
                                if let (Some(e1), Some(e2)) = (r_exprs.next(), r_exprs.next()) {
                                    cond_cases.push(CaseCondition::Range {
                                        from: walk_expr(e1),
                                        to: walk_expr(e2),
                                    });
                                }
                            }
                            Rule::comparison_condition => {
                                let op_str = csub.as_str();
                                let mut c_inner = csub.into_inner();
                                let comp_op = if op_str.starts_with(">=") {
                                    ComparisonOp::GtEq
                                } else if op_str.starts_with("<=") {
                                    ComparisonOp::LtEq
                                } else if op_str.starts_with('>') {
                                    ComparisonOp::Gt
                                } else if op_str.starts_with('<') {
                                    ComparisonOp::Lt
                                } else if op_str.starts_with("!=") {
                                    ComparisonOp::NotEq
                                } else {
                                    ComparisonOp::Eq
                                };
                                if let Some(e) = c_inner.next() {
                                    cond_cases.push(CaseCondition::Comparison {
                                        op: comp_op,
                                        expr: walk_expr(e),
                                    });
                                }
                            }
                            Rule::expr => {
                                cond_cases.push(CaseCondition::Value(walk_expr(csub)));
                            }
                            _ => {}
                        }
                    }
                }
                Rule::block => body_stmts = walk_block_statements(p),
                Rule::statement => {
                    if let Some(s) = walk_statement(p) {
                        body_stmts.push(s);
                    }
                }
                _ => {}
            }
        }

        if is_else {
            default = Some(body_stmts);
        } else if !cond_cases.is_empty() {
            cases.push(SwitchCase {
                conditions: cond_cases,
                body: body_stmts,
            });
        }
    }

    let discriminator = if predicate_mode {
        Expression::bool(true)
    } else {
        subject
    };
    Some(Statement::new(StmtKind::Switch {
        expr: discriminator,
        cases,
        default,
    }))
}

fn walk_for_stmt(pair: Pair<Rule>) -> Option<Statement> {
    let mut var_id = String::new();
    let mut destruct_names = Vec::new();
    let mut iter_expr = Expression::null();
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier => var_id = p.as_str().to_string(),
            Rule::for_destructure => {
                for dsub in p.into_inner() {
                    if dsub.as_rule() == Rule::identifier {
                        destruct_names.push(dsub.as_str().to_string());
                    }
                }
            }
            Rule::expr => iter_expr = walk_expr(p),
            Rule::block => body = walk_block_statements(p),
            Rule::statement => {
                if let Some(s) = walk_statement(p) {
                    body = vec![s];
                }
            }
            _ => {}
        }
    }

    if !destruct_names.is_empty() {
        let loop_tmp = gen_tmp_name();
        let mut prepended_stmts = Vec::new();
        for (idx, name) in destruct_names.clone().into_iter().enumerate() {
            let read_expr = Expression::new(ExprKind::Index {
                object: Box::new(Expression::ident(&loop_tmp)),
                index: Box::new(Expression::int(idx as i64)),
                null_safe: false,
            });
            prepended_stmts.push(Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(name),
                    type_hint: None,
                    init: Some(read_expr),
                    array_bounds: None,
                    with_events: false,
                }],
                kind: VarDeclKind::Const,
            }));
        }
        prepended_stmts.extend(body);
        body = prepended_stmts;
        var_id = loop_tmp;
    }

    let final_iter = if !destruct_names.is_empty() {
        Expression::new(ExprKind::Ternary {
            cond: Box::new(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__coll_is_array")),
                args: vec![Argument::positional(iter_expr.clone())],
                optional: false,
            })),
            then: Box::new(iter_expr.clone()),
            else_: Box::new(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__dict_items")),
                args: vec![Argument::positional(iter_expr)],
                optional: false,
            })),
        })
    } else {
        iter_expr
    };

    Some(Statement::new(StmtKind::ForIn {
        var: var_id,
        key: None,
        iter: final_iter,
        body,
        of: true,
        else_body: None,
        is_async: false,
    }))
}

fn walk_while_stmt(pair: Pair<Rule>) -> Option<Statement> {
    let mut cond = Expression::null();
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::expr => cond = walk_expr(p),
            Rule::block => body = walk_block_statements(p),
            Rule::statement => {
                if let Some(s) = walk_statement(p) {
                    body = vec![s];
                }
            }
            _ => {}
        }
    }

    Some(Statement::new(StmtKind::While {
        cond,
        body,
        else_body: None,
    }))
}

fn walk_do_while_stmt(pair: Pair<Rule>) -> Option<Statement> {
    let mut cond = Expression::null();
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::expr => cond = walk_expr(p),
            Rule::block => body = walk_block_statements(p),
            Rule::statement => {
                if let Some(s) = walk_statement(p) {
                    body = vec![s];
                }
            }
            _ => {}
        }
    }

    Some(Statement::new(StmtKind::DoWhile {
        body,
        cond,
        until: false,
    }))
}

fn walk_lambda(pair: Pair<Rule>) -> Expression {
    let mut params = Vec::new();
    let mut body = Vec::new();
    let mut prefix_stmts = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::lambda_params => {
                for lp in inner.into_inner() {
                    if lp.as_rule() == Rule::lambda_param {
                        let mut name = String::new();
                        let mut destruct_names = Vec::new();
                        let mut type_hint = None;
                        let mut lambda_param_nullable = false;
                        for lsub in lp.into_inner() {
                            match lsub.as_rule() {
                                Rule::identifier => name = lsub.as_str().to_string(),
                                Rule::lambda_destructure => {
                                    for sub in lsub.into_inner() {
                                        if sub.as_rule() == Rule::identifier {
                                            destruct_names.push(sub.as_str().to_string());
                                        }
                                    }
                                }
                                Rule::type_ref => {
                                    lambda_param_nullable = type_ref_is_nullable(lsub.as_str());
                                    type_hint = Some(type_hint_text(lsub.as_str()));
                                }
                                _ => {}
                            }
                        }
                        if !destruct_names.is_empty() {
                            let tmp_param = gen_tmp_name();
                            params.push(Param {
                                name: tmp_param.clone(),
                                type_hint: None,
                                default: None,
                                pass_by: PassBy::Value,
                                is_rest: false,
                                is_kwargs: false,
                                is_optional: false,
                                is_nullable: false,
                            });
                            for (idx, dname) in destruct_names.into_iter().enumerate() {
                                prefix_stmts.push(Statement::new(StmtKind::VarDecl {
                                    declarations: vec![VarDeclarator {
                                        pattern: BindingPattern::Ident(dname),
                                        type_hint: None,
                                        init: Some(Expression::new(ExprKind::Index {
                                            object: Box::new(Expression::ident(&tmp_param)),
                                            index: Box::new(Expression::int(idx as i64)),
                                            null_safe: false,
                                        })),
                                        array_bounds: None,
                                        with_events: false,
                                    }],
                                    kind: VarDeclKind::Const,
                                }));
                            }
                        } else if !name.is_empty() {
                            params.push(Param {
                                name,
                                type_hint,
                                default: None,
                                pass_by: PassBy::Value,
                                is_rest: false,
                                is_kwargs: false,
                                is_optional: false,
                                is_nullable: lambda_param_nullable,
                            });
                        }
                    }
                }
                // prefix_stmts stored in outer scope
            }
            Rule::statement => {
                if let Some(s) = walk_statement(inner) {
                    body.push(s);
                }
            }
            _ => {}
        }
    }

    if !prefix_stmts.is_empty() {
        prefix_stmts.extend(body);
        body = prefix_stmts;
    }

    if params.is_empty() {
        params.push(Param {
            name: "it".to_string(),
            type_hint: None,
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        });
    }

    if let Some(last) = body.pop() {
        match last.kind {
            StmtKind::Expr(e) => {
                body.push(Statement::new(StmtKind::Return(Some(e))));
            }
            other => {
                body.push(Statement::new(other));
            }
        }
    }

    Expression::new(ExprKind::Lambda {
        params,
        body: LambdaBody::Block(body),
        captures: vec![],
        is_async: false,
    })
}

fn walk_expr(pair: Pair<Rule>) -> Expression {
    let rule = pair.as_rule();
    match rule {
        Rule::expr | Rule::assignment => {
            let mut inner = pair.into_inner();
            let first = inner.next().unwrap();
            if let Some(op_pair) = inner.next() {
                let val_pair = inner.next().unwrap();
                let lhs = walk_expr(first);
                let rhs = walk_expr(val_pair);
                let op_str = op_pair.as_str();
                if op_str == "=" {
                    Expression::new(ExprKind::Assign {
                        target: Box::new(lhs),
                        value: Box::new(rhs),
                    })
                } else {
                    let bin_op = match op_str {
                        "+=" => {
                            if matches!(
                                rhs.kind,
                                ExprKind::Binary {
                                    op: BinOp::Concat,
                                    ..
                                } | ExprKind::Lit(Literal::Str(_))
                            ) || matches!(lhs.kind, ExprKind::Lit(Literal::Str(_)))
                            {
                                BinOp::Concat
                            } else {
                                BinOp::Add
                            }
                        }
                        "-=" => BinOp::Sub,
                        "*=" => BinOp::Mul,
                        "/=" => BinOp::Div,
                        "%=" => BinOp::Mod,
                        _ => BinOp::Add,
                    };
                    Expression::new(ExprKind::Assign {
                        target: Box::new(lhs.clone()),
                        value: Box::new(Expression::new(ExprKind::Binary {
                            op: bin_op,
                            left: Box::new(lhs),
                            right: Box::new(rhs),
                        })),
                    })
                }
            } else {
                walk_expr(first)
            }
        }
        Rule::elvis => {
            let mut inner = pair.into_inner();
            let mut current = walk_expr(inner.next().unwrap());
            while let Some(_op) = inner.next() {
                let next_expr = walk_expr(inner.next().unwrap());
                current = Expression::new(ExprKind::NullCoalesce {
                    left: Box::new(current),
                    right: Box::new(next_expr),
                });
            }
            current
        }
        Rule::logical_or => walk_binary_chain(pair, BinOp::Or),
        Rule::logical_and => walk_binary_chain(pair, BinOp::And),
        Rule::equality => {
            let mut inner = pair.into_inner();
            let mut current = walk_expr(inner.next().unwrap());
            while let Some(op_pair) = inner.next() {
                let next_expr = walk_expr(inner.next().unwrap());
                let op = match op_pair.as_str() {
                    "==" => BinOp::Eq,
                    "!=" => BinOp::NotEq,
                    "===" => BinOp::StrictEq,
                    "!==" => BinOp::StrictNotEq,
                    _ => BinOp::Eq,
                };
                current = Expression::new(ExprKind::Binary {
                    op,
                    left: Box::new(current),
                    right: Box::new(next_expr),
                });
            }
            current
        }
        Rule::comparison => {
            let mut inner = pair.into_inner();
            let mut current = walk_expr(inner.next().unwrap());
            while let Some(op_pair) = inner.next() {
                let next_pair = inner.next().unwrap();
                let op_str = op_pair.as_str();
                let type_str = next_pair.as_str().to_lowercase();
                current = match op_str {
                    "<" => Expression::new(ExprKind::Binary {
                        op: BinOp::Lt,
                        left: Box::new(current),
                        right: Box::new(walk_expr(next_pair)),
                    }),
                    "<=" => Expression::new(ExprKind::Binary {
                        op: BinOp::LtEq,
                        left: Box::new(current),
                        right: Box::new(walk_expr(next_pair)),
                    }),
                    ">" => Expression::new(ExprKind::Binary {
                        op: BinOp::Gt,
                        left: Box::new(current),
                        right: Box::new(walk_expr(next_pair)),
                    }),
                    ">=" => Expression::new(ExprKind::Binary {
                        op: BinOp::GtEq,
                        left: Box::new(current),
                        right: Box::new(walk_expr(next_pair)),
                    }),
                    "in" => Expression::new(ExprKind::Binary {
                        op: BinOp::In,
                        left: Box::new(current),
                        right: Box::new(walk_expr(next_pair)),
                    }),
                    "!in" => Expression::new(ExprKind::Unary {
                        op: UnaryOp::Not,
                        expr: Box::new(Expression::new(ExprKind::Binary {
                            op: BinOp::In,
                            left: Box::new(current),
                            right: Box::new(walk_expr(next_pair)),
                        })),
                    }),
                    "is" => Expression::new(ExprKind::IsType {
                        expr: Box::new(current),
                        type_name: type_str,
                    }),
                    "!is" => Expression::new(ExprKind::Unary {
                        op: UnaryOp::Not,
                        expr: Box::new(Expression::new(ExprKind::IsType {
                            expr: Box::new(current),
                            type_name: type_str,
                        })),
                    }),
                    _ => Expression::new(ExprKind::Binary {
                        op: BinOp::Eq,
                        left: Box::new(current),
                        right: Box::new(walk_expr(next_pair)),
                    }),
                };
            }
            current
        }
        Rule::range_expr => {
            let mut inner = pair.into_inner();
            let first = walk_expr(inner.next().unwrap());
            if let Some(_op) = inner.next() {
                let second = walk_expr(inner.next().unwrap());
                Expression::new(ExprKind::Range {
                    start: Box::new(first),
                    end: Box::new(second),
                    inclusive: true,
                })
            } else {
                first
            }
        }
        Rule::additive => {
            let mut inner = pair.into_inner();
            let mut current = walk_expr(inner.next().unwrap());
            while let Some(op_pair) = inner.next() {
                let next_expr = walk_expr(inner.next().unwrap());
                let op = match op_pair.as_str() {
                    // `"a" + x` resolves to `kotlin.String.plus` — CONCATENATION
                    // for every right operand, whatever its type. Emitting a
                    // generic `Add` left the decision to whatever the shared
                    // type inference could see, and an operand it could not
                    // classify (a member read on a user object, a call result)
                    // was coerced toward a number: `"x=" + this.n` trapped in
                    // `toF64` even though `n` held a string. Left-associativity
                    // carries the answer along a chain, so testing the LEFT
                    // operand covers `"a" + x + y`.
                    "+" if kt_is_string_expr(&current) => BinOp::Concat,
                    "+" => BinOp::Add,
                    "-" => BinOp::Sub,
                    _ => BinOp::Add,
                };
                current = Expression::new(ExprKind::Binary {
                    op,
                    left: Box::new(current),
                    right: Box::new(next_expr),
                });
            }
            current
        }
        Rule::multiplicative => {
            let mut inner = pair.into_inner();
            let mut current = walk_expr(inner.next().unwrap());
            while let Some(op_pair) = inner.next() {
                let next_expr = walk_expr(inner.next().unwrap());
                let op_str = op_pair.as_str();
                // `/` used to be rewritten to `(a / b) | 0` here — the JS
                // integer-truncation idiom, applied UNCONDITIONALLY. Kotlin
                // truncates only when BOTH operands are integers, so this made
                // `7.0 / 2.0` answer 3. The shared emitter decides now, from
                // `integer_division_on_slash` plus this language's
                // `[builtin_types] int` spellings (builtinslotplan.md §3i).
                let op = match op_str {
                    "*" => BinOp::Mul,
                    "/" => BinOp::Div,
                    "%" => BinOp::Mod,
                    _ => BinOp::Mul,
                };
                current = Expression::new(ExprKind::Binary {
                    op,
                    left: Box::new(current),
                    right: Box::new(next_expr),
                });
            }
            current
        }
        Rule::type_cast => {
            let mut inner = pair.into_inner();
            let mut current = walk_expr(inner.next().unwrap());
            while let Some(_op_pair) = inner.next() {
                let target_type = inner.next().unwrap().as_str().to_string();
                current = Expression::new(ExprKind::Cast {
                    expr: Box::new(current),
                    type_name: target_type,
                });
            }
            current
        }
        Rule::infix_expr => {
            let mut inner = pair.into_inner();
            let mut current = walk_expr(inner.next().unwrap());
            while let Some(op_pair) = inner.next() {
                let next_expr = walk_expr(inner.next().unwrap());
                let op_str = op_pair.as_str();
                if op_str == "to" {
                    // Kotlin `a to b` → Pair(a, b)
                    current = create_pair_expr(current, next_expr);
                } else if op_str == "until" {
                    // a until b → exclusive ascending [a, a+1, ..., b-1]
                    current = Expression::new(ExprKind::Range {
                        start: Box::new(current),
                        end: Box::new(next_expr),
                        inclusive: false,
                    });
                } else if op_str == "downTo" {
                    // a downTo b → descending [a, a-1, ..., b], INCLUSIVE of b.
                    // Maps to `__kt_step_desc(a, b - 1, -1)` → `collections.range_step`,
                    // whose stop is EXCLUSIVE (it iterates while `i > stop` for a
                    // negative step). Passing `b` straight through dropped the last
                    // element: `5 downTo 2` yielded 5,4,3.
                    current = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__kt_step_desc")),
                        args: vec![
                            Argument::positional(current),
                            Argument::positional(Expression::new(ExprKind::Binary {
                                op: BinOp::Sub,
                                left: Box::new(next_expr),
                                right: Box::new(Expression::int(1)),
                            })),
                            Argument::positional(Expression::new(ExprKind::Unary {
                                op: UnaryOp::Neg,
                                expr: Box::new(Expression::int(1)),
                            })),
                        ],
                        optional: false,
                    });
                } else if op_str == "step" {
                    // `range step n` — must convert range to 3-arg stepped form.
                    match current.kind.clone() {
                        // (a downTo b) step n  → replace -1 with -n
                        ExprKind::Call {
                            callee,
                            mut args,
                            optional,
                        } if matches!(&callee.kind, ExprKind::Ident(nm) if nm == "__kt_step_desc") =>
                        {
                            if args.len() == 3 {
                                args[2] = Argument::positional(Expression::new(ExprKind::Unary {
                                    op: UnaryOp::Neg,
                                    expr: Box::new(next_expr),
                                }));
                            }
                            current = Expression::new(ExprKind::Call {
                                callee,
                                args,
                                optional,
                            });
                        }
                        // (a..b) step n  or  (a until b) step n
                        ExprKind::Range {
                            start,
                            end,
                            inclusive,
                        } => {
                            let stop = if inclusive {
                                // inclusive end+1 so the 3-arg exclusive loop includes end
                                Expression::new(ExprKind::Binary {
                                    op: BinOp::Add,
                                    left: end,
                                    right: Box::new(Expression::int(1)),
                                })
                            } else {
                                *end
                            };
                            current = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::ident("__kt_step_asc")),
                                args: vec![
                                    Argument::positional(*start),
                                    Argument::positional(stop),
                                    Argument::positional(next_expr),
                                ],
                                optional: false,
                            });
                        }
                        _ => {
                            // Fallback: pass through as method call
                            current = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::new(ExprKind::Member {
                                    object: Box::new(current),
                                    field: "step".to_string(),
                                    null_safe: false,
                                })),
                                args: vec![Argument::positional(next_expr)],
                                optional: false,
                            });
                        }
                    }
                } else if let Some(op) = infix_bitwise_op(op_str) {
                    // Kotlin spells the bitwise operators as infix functions
                    // (`6 and 3`, `1 shl 2`). They are the SAME operators every
                    // other language writes with punctuation, so they lower to
                    // the shared `BinOp` and reach `primitives/operators.rs`
                    // rather than becoming an `Int.and(…)` member call that no
                    // primitive implements.
                    current = Expression::new(ExprKind::Binary {
                        op,
                        left: Box::new(current),
                        right: Box::new(next_expr),
                    });
                } else {
                    current = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Member {
                            object: Box::new(current),
                            field: op_str.to_string(),
                            null_safe: false,
                        })),
                        args: vec![Argument::positional(next_expr)],
                        optional: false,
                    });
                }
            }
            current
        }
        Rule::unary => {
            let mut inner = pair.into_inner();
            let mut ops = Vec::new();
            while let Some(p) = inner.next() {
                if p.as_rule() == Rule::prefix_op {
                    ops.push(p.as_str().to_string());
                } else {
                    let mut current = walk_expr(p);
                    for op in ops.into_iter().rev() {
                        let un_op = match op.as_str() {
                            "!" => UnaryOp::Not,
                            "-" => UnaryOp::Neg,
                            "+" => UnaryOp::Pos,
                            _ => UnaryOp::Not,
                        };
                        current = Expression::new(ExprKind::Unary {
                            op: un_op,
                            expr: Box::new(current),
                        });
                    }
                    return current;
                }
            }
            Expression::null()
        }
        Rule::postfix | Rule::delegate_expr => {
            let parts: Vec<_> = pair.into_inner().collect();
            let primary_pair = parts[0].clone();
            // `super<Base>` names WHICH supertype answers. The grammar has
            // carried the qualifier since generics landed; the walker dropped
            // it, so `super<A>.tag()` and `super<I>.tag()` compiled the same.
            let super_qualifier: Option<String> = {
                fn find_super_type(pair: Pair<Rule>) -> Option<String> {
                    if pair.as_rule() == Rule::super_expr {
                        return pair
                            .into_inner()
                            .find(|p| p.as_rule() == Rule::type_ref)
                            .map(|p| p.as_str().trim().to_string());
                    }
                    pair.into_inner().find_map(find_super_type)
                }
                find_super_type(primary_pair.clone())
            };
            let mut current = walk_expr(primary_pair);

            for (idx, suffix_pair) in parts.iter().skip(1).enumerate() {
                // Whether a call follows THIS suffix: `super.v` is a property
                // read, `super.v()` is a call, and only the second is a
                // `SuperCall`.
                let next_is_call = parts
                    .get(idx + 2)
                    .and_then(|p| p.clone().into_inner().next())
                    .map(|q| q.as_rule() == Rule::call_suffix)
                    .unwrap_or(false);
                let suffix_pair = suffix_pair.clone();
                let suffix_inner = suffix_pair.into_inner().next().unwrap();
                match suffix_inner.as_rule() {
                    Rule::type_args => {
                        continue;
                    }
                    Rule::call_suffix => {
                        let mut args = Vec::new();
                        for item in suffix_inner.into_inner() {
                            match item.as_rule() {
                                Rule::arg_list => {
                                    for arg_p in item.into_inner() {
                                        let mut arg_expr = None;
                                        let mut arg_name = None;
                                        let mut is_spread = false;
                                        for sub in arg_p.into_inner() {
                                            match sub.as_rule() {
                                                Rule::spread_op => is_spread = true,
                                                Rule::identifier => {
                                                    if arg_name.is_none() && arg_expr.is_none() {
                                                        arg_name = Some(sub.as_str().to_string());
                                                    }
                                                }
                                                Rule::expr => arg_expr = Some(walk_expr(sub)),
                                                _ => {}
                                            }
                                        }
                                        if let Some(ae) = arg_expr {
                                            args.push(Argument {
                                                value: ae,
                                                name: arg_name,
                                                by_ref: false,
                                                spread: is_spread,
                                            });
                                        }
                                    }
                                }
                                Rule::lambda_literal => {
                                    args.push(Argument::positional(walk_lambda(item)));
                                }
                                _ => {}
                            }
                        }

                        if let ExprKind::Member {
                            ref object,
                            ref field,
                            null_safe: _,
                        } = current.clone().kind
                        {
                            match field.as_str() {
                                "put" if args.len() == 2 => {
                                    current = Expression::new(ExprKind::Assign {
                                        target: Box::new(Expression::new(ExprKind::Index {
                                            object: object.clone(),
                                            index: Box::new(args[0].value.clone()),
                                            null_safe: false,
                                        })),
                                        value: Box::new(args[1].value.clone()),
                                    });
                                    continue;
                                }
                                "get" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Index {
                                        object: object.clone(),
                                        index: Box::new(args[0].value.clone()),
                                        null_safe: false,
                                    });
                                    continue;
                                }
                                "getOrDefault" if args.len() == 2 => {
                                    current = Expression::new(ExprKind::NullCoalesce {
                                        left: Box::new(Expression::new(ExprKind::Index {
                                            object: object.clone(),
                                            index: Box::new(args[0].value.clone()),
                                            null_safe: false,
                                        })),
                                        right: Box::new(args[1].value.clone()),
                                    });
                                    continue;
                                }
                                "containsKey" | "contains" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Binary {
                                        op: BinOp::In,
                                        left: Box::new(args[0].value.clone()),
                                        right: object.clone(),
                                    });
                                    continue;
                                }
                                // NOTE: `.add(x)` for Set (dict) is handled in the second
                                // Member block below via __coll_push, which works uniformly
                                // for both list (array.push) and set (set semantics via
                                // array.push on the keys array). Do NOT intercept it here.
                                "remove" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Delete(Box::new(
                                        Expression::new(ExprKind::Index {
                                            object: object.clone(),
                                            index: Box::new(args[0].value.clone()),
                                            null_safe: false,
                                        }),
                                    )));
                                    continue;
                                }
                                "clear" if args.is_empty() => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__dict_clear")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "isEmpty" if args.is_empty() => {
                                    let keys_len = Expression::new(ExprKind::Member {
                                        object: object.clone(),
                                        field: "length".to_string(),
                                        null_safe: false,
                                    });
                                    current = Expression::new(ExprKind::Binary {
                                        op: BinOp::Eq,
                                        left: Box::new(keys_len),
                                        right: Box::new(Expression::int(0)),
                                    });
                                    continue;
                                }
                                "isNotEmpty" if args.is_empty() => {
                                    let keys_len = Expression::new(ExprKind::Member {
                                        object: object.clone(),
                                        field: "length".to_string(),
                                        null_safe: false,
                                    });
                                    current = Expression::new(ExprKind::Binary {
                                        op: BinOp::Gt,
                                        left: Box::new(keys_len),
                                        right: Box::new(Expression::int(0)),
                                    });
                                    continue;
                                }
                                _ => {}
                            }
                        }

                        let func_name = match &current.kind {
                            ExprKind::Ident(name) => Some(name.clone()),
                            _ => None,
                        };

                        if let Some(ref fn_name) = func_name {
                            if fn_name == "Pair" && args.len() == 2 && !is_user_class_name(fn_name)
                            {
                                current =
                                    create_pair_expr(args[0].value.clone(), args[1].value.clone());
                                continue;
                            }
                            if fn_name == "Triple"
                                && args.len() == 3
                                && !is_user_class_name(fn_name)
                            {
                                current = create_triple_expr(
                                    args[0].value.clone(),
                                    args[1].value.clone(),
                                    args[2].value.clone(),
                                );
                                continue;
                            }
                            if matches!(
                                fn_name.as_str(),
                                "mapOf"
                                    | "mutableMapOf"
                                    | "linkedMapOf"
                                    | "hashMapOf"
                                    | "buildMap"
                                    | "emptyMap"
                            ) {
                                let mut props = Vec::new();
                                for arg in args {
                                    if let ExprKind::Object(ref pair_props) = arg.value.kind {
                                        let mut k_expr = None;
                                        let mut v_expr = None;
                                        for p in pair_props {
                                            if let ObjectProperty::KeyValue { key, value } = p {
                                                if let ExprKind::Lit(Literal::Str(s)) = &key.kind {
                                                    if s == "0" || s == "first" {
                                                        if k_expr.is_none() {
                                                            k_expr = Some(value.clone());
                                                        }
                                                    } else if s == "1" || s == "second" {
                                                        if v_expr.is_none() {
                                                            v_expr = Some(value.clone());
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        if let (Some(k), Some(v)) = (k_expr, v_expr) {
                                            props.push(ObjectProperty::KeyValue {
                                                key: k,
                                                value: v,
                                            });
                                            continue;
                                        }
                                    }
                                    if let ExprKind::Tuple(ref pair_elems) = arg.value.kind {
                                        if pair_elems.len() == 2 {
                                            props.push(ObjectProperty::KeyValue {
                                                key: pair_elems[0].clone(),
                                                value: pair_elems[1].clone(),
                                            });
                                            continue;
                                        }
                                    }
                                    if let ExprKind::Array(ref pair_elems) = arg.value.kind {
                                        if pair_elems.len() == 2 {
                                            props.push(ObjectProperty::KeyValue {
                                                key: pair_elems[0].value.clone(),
                                                value: pair_elems[1].value.clone(),
                                            });
                                            continue;
                                        }
                                    }
                                    props.push(ObjectProperty::KeyValue {
                                        key: Expression::new(ExprKind::Index {
                                            object: Box::new(arg.value.clone()),
                                            index: Box::new(Expression::int(0)),
                                            null_safe: false,
                                        }),
                                        value: Expression::new(ExprKind::Index {
                                            object: Box::new(arg.value.clone()),
                                            index: Box::new(Expression::int(1)),
                                            null_safe: false,
                                        }),
                                    });
                                }
                                current = create_map_expr(props);
                                continue;
                            }
                            if matches!(
                                fn_name.as_str(),
                                "setOf"
                                    | "mutableSetOf"
                                    | "linkedSetOf"
                                    | "hashSetOf"
                                    | "buildSet"
                                    | "emptySet"
                            ) {
                                let elems = args.into_iter().map(|a| a.value).collect();
                                current = create_kotlin_set_expr(elems);
                                continue;
                            }
                            if fn_name == "joinToString" && !args.is_empty() {
                                let separator =
                                    args.get(1).map(|arg| arg.value.clone()).unwrap_or_else(|| {
                                        Expression::new(ExprKind::Lit(Literal::Str(
                                            ", ".to_string(),
                                        )))
                                    });
                                current = Expression::new(ExprKind::Call {
                                    callee: Box::new(Expression::ident("__coll_join")),
                                    args: vec![
                                        Argument::positional(args[0].value.clone()),
                                        Argument::positional(separator),
                                    ],
                                    optional: false,
                                });
                                continue;
                            }
                            if matches!(
                                fn_name.as_str(),
                                "listOf"
                                    | "mutableListOf"
                                    | "arrayOf"
                                    | "emptyList"
                                    | "intArrayOf"
                                    | "doubleArrayOf"
                                    | "booleanArrayOf"
                                    | "charArrayOf"
                                    | "longArrayOf"
                                    | "sequenceOf"
                            ) {
                                let elements = args
                                    .into_iter()
                                    .map(|a| ArrayElement {
                                        key: None,
                                        value: a.value,
                                        spread: false,
                                        by_ref: false,
                                    })
                                    .collect();
                                current = Expression::new(ExprKind::Array(elements));
                                continue;
                            }
                        }

                        // `x.ext(a)` where `ext` is a top-level `fun X.ext(a)`
                        // — an extension is a FUNCTION whose first parameter is
                        // the receiver, so the receiver moves into the argument
                        // list. A real member of the same name wins, which is
                        // Kotlin's rule.
                        if let ExprKind::Member {
                            ref object,
                            ref field,
                            ..
                        } = current.kind
                        {
                            if is_extension_function(field)
                                && !is_user_member_name(field, args.len())
                            {
                                let mut ext_args = vec![Argument::positional(*object.clone())];
                                ext_args.extend(args);
                                current = Expression::new(ExprKind::Call {
                                    callee: Box::new(Expression::ident(field)),
                                    args: ext_args,
                                    optional: false,
                                });
                                continue;
                            }
                        }

                        if let ExprKind::Member { ref mut field, .. } = current.kind {
                            if let Some(storage_name) = overloaded_storage_name(field, args.len()) {
                                *field = storage_name;
                            }
                        }

                        // A member some class in this source DECLARES is never
                        // rewritten to a collection primitive. The rewrites
                        // below match on SPELLING and cannot see the receiver's
                        // type, so `class Calc { fun add(x: Int) = base + x }`
                        // had `c.add(3)` compiled as an array push, answering
                        // `1`. Twenty collection names were stolen this way from
                        // every object in the program. A declared member wins,
                        // which is Kotlin's own rule.
                        let rewritable = match current.kind {
                            ExprKind::Member { ref field, .. } => {
                                !is_user_member_name(field, args.len())
                            }
                            _ => false,
                        };
                        if let ExprKind::Member {
                            ref object,
                            ref field,
                            ..
                        } = current.kind
                            && rewritable
                        {
                            match field.as_str() {
                                "put" if args.len() == 2 => {
                                    current = Expression::new(ExprKind::Assign {
                                        target: Box::new(Expression::new(ExprKind::Index {
                                            object: object.clone(),
                                            index: Box::new(args[0].value.clone()),
                                            null_safe: false,
                                        })),
                                        value: Box::new(args[1].value.clone()),
                                    });
                                    continue;
                                }
                                "get" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Index {
                                        object: object.clone(),
                                        index: Box::new(args[0].value.clone()),
                                        null_safe: false,
                                    });
                                    continue;
                                }
                                "getOrDefault" if args.len() == 2 => {
                                    let get_expr = Expression::new(ExprKind::Index {
                                        object: object.clone(),
                                        index: Box::new(args[0].value.clone()),
                                        null_safe: false,
                                    });
                                    current = Expression::new(ExprKind::NullCoalesce {
                                        left: Box::new(get_expr),
                                        right: Box::new(args[1].value.clone()),
                                    });
                                    continue;
                                }
                                "containsKey" | "contains" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__dict_has")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            Argument::positional(args[0].value.clone()),
                                        ],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "containsValue" if args.len() == 1 => {
                                    let values_expr = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__dict_values")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false,
                                    });
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_contains")),
                                        args: vec![
                                            Argument::positional(values_expr),
                                            Argument::positional(args[0].value.clone()),
                                        ],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "isEmpty" if args.is_empty() => {
                                    let sz = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__dict_size")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false,
                                    });
                                    current = Expression::new(ExprKind::Binary {
                                        op: BinOp::Eq,
                                        left: Box::new(sz),
                                        right: Box::new(Expression::int(0)),
                                    });
                                    continue;
                                }
                                "isNotEmpty" if args.is_empty() => {
                                    let sz = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__dict_size")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false,
                                    });
                                    current = Expression::new(ExprKind::Binary {
                                        op: BinOp::Gt,
                                        left: Box::new(sz),
                                        right: Box::new(Expression::int(0)),
                                    });
                                    continue;
                                }
                                "remove" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Delete(Box::new(
                                        Expression::new(ExprKind::Index {
                                            object: object.clone(),
                                            index: Box::new(args[0].value.clone()),
                                            null_safe: false,
                                        }),
                                    )));
                                    continue;
                                }
                                "removeAt" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_removeAt")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            Argument::positional(args[0].value.clone()),
                                        ],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "clear" if args.is_empty() => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__dict_clear")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "add" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_push")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            Argument::positional(args[0].value.clone()),
                                        ],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "add" if args.len() == 2 => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_insert")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            Argument::positional(args[0].value.clone()),
                                            Argument::positional(args[1].value.clone()),
                                        ],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "indexOf" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_indexOf")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            Argument::positional(args[0].value.clone()),
                                        ],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "lastIndexOf" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_lastIndexOf")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            Argument::positional(args[0].value.clone()),
                                        ],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "reversed" if args.is_empty() => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_reverse")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "sorted" if args.is_empty() => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_sorted")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "joinToString" if args.len() >= 2 => {
                                    let mapped = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__array_map")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[1].clone(),
                                        ],
                                        optional: false,
                                    });
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_join")),
                                        args: vec![
                                            Argument::positional(mapped),
                                            Argument::positional(args[0].value.clone()),
                                        ],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "joinToString" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_join")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            Argument::positional(args[0].value.clone()),
                                        ],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "joinToString" if args.is_empty() => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_join")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            Argument::positional(Expression::new(ExprKind::Lit(
                                                Literal::Str(", ".to_string()),
                                            ))),
                                        ],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "sortedBy" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::new(ExprKind::Member {
                                            object: object.clone(),
                                            field: "sort".to_string(),
                                            null_safe: false,
                                        })),
                                        args,
                                        optional: false,
                                    });
                                    continue;
                                }
                                "contentToString" if args.is_empty() => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__kt_tostring")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "sum" if args.is_empty() => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_sum")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "fold" if args.len() == 2 => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::new(ExprKind::Member {
                                            object: object.clone(),
                                            field: "__array_reduce".to_string(),
                                            null_safe: false,
                                        })),
                                        args: vec![
                                            Argument::positional(args[1].value.clone()),
                                            Argument::positional(args[0].value.clone()),
                                        ],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "take" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_slice")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            Argument::positional(Expression::int(0)),
                                            Argument::positional(args[0].value.clone()),
                                        ],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "drop" if args.len() == 1 => {
                                    let len_expr = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_length")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false,
                                    });
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_slice")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            Argument::positional(args[0].value.clone()),
                                            Argument::positional(len_expr),
                                        ],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "first" | "firstOrNull" if args.is_empty() => {
                                    current = Expression::new(ExprKind::Index {
                                        object: object.clone(),
                                        index: Box::new(Expression::int(0)),
                                        null_safe: false,
                                    });
                                    continue;
                                }
                                "max" | "maxOrNull" if args.is_empty() => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_max")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "min" | "minOrNull" if args.is_empty() => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_min")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false,
                                    });
                                    continue;
                                }
                                "filter" | "filterNot" | "map" | "forEach" if args.len() == 1 => {
                                    let iter_target = Expression::new(ExprKind::Ternary {
                                        cond: Box::new(Expression::new(ExprKind::Call {
                                            callee: Box::new(Expression::ident("__coll_is_array")),
                                            args: vec![Argument::positional(*object.clone())],
                                            optional: false,
                                        })),
                                        then: Box::new(*object.clone()),
                                        else_: Box::new(Expression::new(ExprKind::Call {
                                            callee: Box::new(Expression::ident("__dict_items")),
                                            args: vec![Argument::positional(*object.clone())],
                                            optional: false,
                                        })),
                                    });
                                    // Keep the SOURCE spelling. `[array_methods]`
                                    // is keyed by it (`filter = "__array_filter"`),
                                    // and `lookup_array_method` looks up the KEY —
                                    // rewriting the field to `__array_filter` here
                                    // meant the lookup missed, the higher-order
                                    // dispatch never ran, and the call fell through
                                    // to a member named `__array_filter` that does
                                    // not exist ("undefined is not callable").
                                    let emit_method = match field.as_str() {
                                        "filterNot" => "filter",
                                        other => other,
                                    };
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::new(ExprKind::Member {
                                            object: Box::new(iter_target),
                                            field: emit_method.to_string(),
                                            null_safe: false,
                                        })),
                                        args,
                                        optional: false,
                                    });
                                    continue;
                                }
                                _ => {}
                            }
                        }

                        // Kotlin has no `new`, so a call whose callee names a
                        // TYPE is a construction. That is true of a qualified
                        // spelling too — `java.util.ArrayList()`,
                        // `java.math.BigInteger("1")` — which stayed an
                        // ordinary member call and trapped, while the
                        // `import`ed form worked. Same rule, applied to the
                        // last segment of the chain.
                        let is_type_spelling = |name: &str| {
                            name.chars().next().is_some_and(char::is_uppercase)
                                && !matches!(
                                    name,
                                    "Exception"
                                        | "IllegalArgumentException"
                                        | "IllegalStateException"
                                        | "NullPointerException"
                                        | "IndexOutOfBoundsException"
                                )
                        };
                        let is_class_name = match &current.kind {
                            ExprKind::Ident(name) => is_type_spelling(name),
                            ExprKind::Member { field, .. }
                                if qualified_inner_class(field).is_some() =>
                            {
                                true
                            }
                            // Only a chain of plain idents is a qualified type
                            // name; `expr.Foo()` is a method call on a value.
                            ExprKind::Member { object, field, .. } => {
                                is_type_spelling(field) && is_ident_chain(object)
                            }
                            _ => false,
                        };

                        if is_class_name {
                            if let ExprKind::Member { object, field, .. } = &current.kind {
                                if let Some(qualified) = qualified_inner_class(field) {
                                    let mut with_outer =
                                        vec![Argument::positional(*object.clone())];
                                    with_outer.extend(args);
                                    args = with_outer;
                                    current = qualified_type_expr(&qualified);
                                }
                            } else if let ExprKind::Ident(cname) = &current.kind {
                                if let Some(qualified) = qualified_inner_class(cname) {
                                    let owner = qualified
                                        .rsplit_once('.')
                                        .map(|(owner, _)| owner)
                                        .unwrap_or("");
                                    let inside_owner = CURRENT_CLASS_STACK.with(|stack| {
                                        stack.borrow().iter().any(|name| name == owner)
                                    });
                                    if inside_owner {
                                        let mut with_outer = vec![Argument::positional(
                                            Expression::new(ExprKind::This),
                                        )];
                                        with_outer.extend(args);
                                        args = with_outer;
                                        current = qualified_type_expr(&qualified);
                                    }
                                }
                            }
                            // A local class receives what it captured, ahead of
                            // whatever the source passes.
                            if let ExprKind::Ident(ref cname) = current.kind {
                                let caps = LOCAL_CLASS_CAPTURES
                                    .with(|m| m.borrow().get(cname).cloned())
                                    .unwrap_or_default();
                                if !caps.is_empty() {
                                    let mut with_caps: Vec<Argument> = caps
                                        .iter()
                                        .map(|c| Argument::positional(Expression::ident(c)))
                                        .collect();
                                    with_caps.extend(args);
                                    args = with_caps;
                                }
                            }
                            current = Expression::new(ExprKind::New {
                                class: Box::new(current),
                                args,
                            });
                        } else {
                            if let ExprKind::Member {
                                ref object,
                                ref field,
                                ..
                            } = current.kind
                            {
                                if matches!(
                                    field.as_str(),
                                    "filter" | "filterNot" | "map" | "forEach"
                                ) && args.len() == 1
                                {
                                    let iter_target = Expression::new(ExprKind::Ternary {
                                        cond: Box::new(Expression::new(ExprKind::Call {
                                            callee: Box::new(Expression::ident("__coll_is_array")),
                                            args: vec![Argument::positional(*object.clone())],
                                            optional: false,
                                        })),
                                        then: Box::new(*object.clone()),
                                        else_: Box::new(Expression::new(ExprKind::Call {
                                            callee: Box::new(Expression::ident("__dict_items")),
                                            args: vec![Argument::positional(*object.clone())],
                                            optional: false,
                                        })),
                                    });
                                    // Keep the SOURCE spelling. `[array_methods]`
                                    // is keyed by it (`filter = "__array_filter"`),
                                    // and `lookup_array_method` looks up the KEY —
                                    // rewriting the field to `__array_filter` here
                                    // meant the lookup missed, the higher-order
                                    // dispatch never ran, and the call fell through
                                    // to a member named `__array_filter` that does
                                    // not exist ("undefined is not callable").
                                    let emit_method = match field.as_str() {
                                        "filterNot" => "filter",
                                        other => other,
                                    };
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::new(ExprKind::Member {
                                            object: Box::new(iter_target),
                                            field: emit_method.to_string(),
                                            null_safe: false,
                                        })),
                                        args,
                                        optional: false,
                                    });
                                    continue;
                                }
                            }
                            // `super.f(a)` — `member_suffix` already turned
                            // `super.f` into a `SuperCall`, which IS the call.
                            // Wrapping it again called the RESULT ("string is
                            // not callable"), so the arguments land on the node
                            // that already exists instead of a second one.
                            if let ExprKind::SuperCall {
                                args: super_args, ..
                            } = &mut current.kind
                            {
                                if super_args.is_empty() {
                                    *super_args = args;
                                    continue;
                                }
                            }
                            // `.toString` used to be rewritten into `"" + x`
                            // here, so the `()` that follows it had to be
                            // swallowed. That rewrite is gone: `toString` is a
                            // MEMBER, a class may override it, and the built-in
                            // rendering is declared as a value method
                            // (`common:kotlin.tostring`) instead. Rendering
                            // dispatches on the VALUE, in `emitter/tostring.rs`.
                            current = Expression::new(ExprKind::Call {
                                callee: Box::new(current),
                                args,
                                optional: false,
                            });
                        }
                    }
                    Rule::member_suffix => {
                        let field_id = suffix_inner
                            .into_inner()
                            .next()
                            .unwrap()
                            .as_str()
                            .to_string();
                        if let ExprKind::Super = current.kind {
                            let parent = CURRENT_CLASS_PARENT
                                .with(|stack| stack.borrow().last().cloned().flatten());
                            let qualifies_interface = super_qualifier
                                .as_ref()
                                .is_some_and(|q| parent.as_deref() != Some(q.as_str()));
                            if !next_is_call {
                                // A stored `override val` shares the base's slot
                                // and initializers run base-first, so at the
                                // point the override's initializer runs, the
                                // field still holds the BASE's value. That is
                                // exactly what `super.p` names.
                                current = Expression::new(ExprKind::Member {
                                    object: Box::new(Expression::new(ExprKind::This)),
                                    field: field_id,
                                    null_safe: false,
                                });
                            } else if qualifies_interface {
                                // `super<I>.m()` — the class's own `m` shadowed
                                // the interface default, so reach it through the
                                // ALIAS the augmentation binds it under. The
                                // shared fold's `rename_to` is additive, which
                                // is the whole mechanism (§4c).
                                let from = super_qualifier.clone().unwrap();
                                SUPER_QUALIFIED_USES.with(|stack| {
                                    if let Some(top) = stack.borrow_mut().last_mut() {
                                        top.push((from.clone(), field_id.clone()));
                                    }
                                });
                                current = Expression::new(ExprKind::Member {
                                    object: Box::new(Expression::new(ExprKind::This)),
                                    field: super_alias(&from, &field_id),
                                    null_safe: false,
                                });
                            } else {
                                current = Expression::new(ExprKind::SuperCall {
                                    method: Some(field_id),
                                    args: vec![],
                                });
                            }
                        } else if is_extension_property(&field_id)
                            && !is_user_property_name(&field_id)
                        {
                            current = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::ident(&field_id)),
                                args: vec![Argument::positional(current)],
                                optional: false,
                            });
                        } else if is_user_property_name(&field_id) {
                            // A property a class in this source declares is an
                            // ordinary member read. The rewrites below match on
                            // SPELLING, so `data class Counter(val values:
                            // MutableList<Int>)` had `a.values` answer the
                            // OBJECT's members via `__dict_values`.
                            current = Expression::new(ExprKind::Member {
                                object: Box::new(current),
                                field: field_id.clone(),
                                null_safe: false,
                            });
                        } else {
                            match field_id.as_str() {
                                // `first`/`second`/`third` are PROPERTIES on
                                // `Pair`/`Triple`, which lower to an array
                                // (`common:collections.new`), so the positional
                                // read is the whole meaning. `componentN` is
                                // deliberately NOT here: it is a FUNCTION, and a
                                // `data class` declares its own — rewriting the
                                // member turned `u.component1()` into `u[0]()`
                                // ("string is not callable") and made the
                                // synthesized member unreachable. It resolves as
                                // an ordinary member call now, like any other.
                                "first" => {
                                    current = Expression::new(ExprKind::Index {
                                        object: Box::new(current),
                                        index: Box::new(Expression::int(0)),
                                        null_safe: false,
                                    });
                                }
                                "second" => {
                                    current = Expression::new(ExprKind::Index {
                                        object: Box::new(current),
                                        index: Box::new(Expression::int(1)),
                                        null_safe: false,
                                    });
                                }
                                "third" => {
                                    current = Expression::new(ExprKind::Index {
                                        object: Box::new(current),
                                        index: Box::new(Expression::int(2)),
                                        null_safe: false,
                                    });
                                }
                                "keys" => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__dict_keys")),
                                        args: vec![Argument::positional(current)],
                                        optional: false,
                                    });
                                }
                                "values" => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__dict_values")),
                                        args: vec![Argument::positional(current)],
                                        optional: false,
                                    });
                                }
                                "entries" => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__dict_items")),
                                        args: vec![Argument::positional(current)],
                                        optional: false,
                                    });
                                }
                                "size" | "length" => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_length")),
                                        args: vec![Argument::positional(current)],
                                        optional: false,
                                    });
                                }
                                "lastIndex" => {
                                    let len_expr = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_length")),
                                        args: vec![Argument::positional(current)],
                                        optional: false,
                                    });
                                    current = Expression::new(ExprKind::Binary {
                                        op: BinOp::Sub,
                                        left: Box::new(len_expr),
                                        right: Box::new(Expression::int(1)),
                                    });
                                }
                                "indices" => {
                                    let len_expr = Expression::new(ExprKind::Member {
                                        object: Box::new(current.clone()),
                                        field: "length".to_string(),
                                        null_safe: false,
                                    });
                                    current = Expression::new(ExprKind::Range {
                                        start: Box::new(Expression::int(0)),
                                        end: Box::new(Expression::new(ExprKind::Binary {
                                            op: BinOp::Sub,
                                            left: Box::new(len_expr),
                                            right: Box::new(Expression::int(1)),
                                        })),
                                        inclusive: true,
                                    });
                                }
                                _ => {
                                    current = Expression::new(ExprKind::Member {
                                        object: Box::new(current),
                                        field: field_id,
                                        null_safe: false,
                                    });
                                }
                            }
                        }
                    }
                    Rule::safe_call_suffix => {
                        let field_id = suffix_inner
                            .into_inner()
                            .next()
                            .unwrap()
                            .as_str()
                            .to_string();
                        current = Expression::new(ExprKind::Member {
                            object: Box::new(current),
                            field: field_id,
                            null_safe: true,
                        });
                    }
                    Rule::index_suffix => {
                        let index_pair = suffix_inner.into_inner().next().unwrap();
                        let idx_expr = walk_expr(
                            index_pair
                                .into_inner()
                                .next()
                                .unwrap()
                                .into_inner()
                                .next()
                                .unwrap(),
                        );
                        current = Expression::new(ExprKind::Index {
                            object: Box::new(current),
                            index: Box::new(idx_expr),
                            null_safe: false,
                        });
                    }
                    Rule::null_assert_suffix => {
                        current = Expression::new(ExprKind::Unary {
                            op: UnaryOp::Not,
                            expr: Box::new(current),
                        });
                    }
                    Rule::inc_suffix => {
                        let op_str = suffix_inner.as_str();
                        let bin_op = if op_str == "++" {
                            BinOp::Add
                        } else {
                            BinOp::Sub
                        };
                        current = Expression::new(ExprKind::Assign {
                            target: Box::new(current.clone()),
                            value: Box::new(Expression::new(ExprKind::Binary {
                                op: bin_op,
                                left: Box::new(current),
                                right: Box::new(Expression::int(1)),
                            })),
                        });
                    }
                    _ => {}
                }
            }
            current
        }
        Rule::primary => {
            let inner = pair.into_inner().next().unwrap();
            match inner.as_rule() {
                Rule::identifier => {
                    let name = backing_field_substitution(inner.as_str());
                    inner_outer_read(&name)
                        .or_else(|| extension_receiver_read(&name))
                        .unwrap_or_else(|| Expression::ident(&name))
                }
                Rule::callable_ref => walk_callable_ref(inner),
                Rule::literal => walk_literal(inner),
                Rule::this_kw => Expression::new(ExprKind::This),
                Rule::super_kw => Expression::new(ExprKind::Super),
                // `this@Outer` / `super<Base>` / `super@Outer`. The label and
                // the explicit supertype are RESOLUTION hints; the receiver
                // itself is still `this` / `super`, so the concept node is the
                // same one an unqualified occurrence produces and no downstream
                // path has to learn a second shape.
                Rule::this_expr => {
                    let text = inner.as_str();
                    if let Some(label) = text.split_once('@').map(|(_, label)| label.trim()) {
                        INNER_OUTER_MEMBERS.with(|stack| {
                            stack
                                .borrow()
                                .last()
                                .and_then(|outers| {
                                    outers
                                        .iter()
                                        .position(|(name, _)| name == label)
                                        .map(|idx| outer_this(idx + 1))
                                })
                                .unwrap_or_else(|| Expression::new(ExprKind::This))
                        })
                    } else {
                        Expression::new(ExprKind::This)
                    }
                }
                Rule::super_expr => Expression::new(ExprKind::Super),
                Rule::lambda_literal => walk_lambda(inner),
                Rule::object_expr => {
                    let mut parent = None;
                    let mut interfaces = Vec::new();
                    let mut members = Vec::new();
                    for osub in inner.into_inner() {
                        match osub.as_rule() {
                            Rule::inheritance_list => {
                                for spec in osub.into_inner() {
                                    if spec.as_rule() == Rule::inheritance_specifier {
                                        // Parentheses mark the SUPERCLASS, the
                                        // same rule `walk_class_decl` uses:
                                        // `object : Base(), I` extends `Base`.
                                        // Taking the first supertype made
                                        // `object : Callback { … }` extend an
                                        // interface it should have implemented.
                                        let calls_constructor = spec.as_str().contains('(');
                                        for sub in spec.into_inner() {
                                            if sub.as_rule() == Rule::type_ref {
                                                let tname = sub.as_str().trim().to_string();
                                                if calls_constructor && parent.is_none() {
                                                    parent =
                                                        Some(Box::new(Expression::ident(&tname)));
                                                } else {
                                                    interfaces.push(tname);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            Rule::class_body => {
                                let mut body_members = walk_object_body_members(osub, false);
                                // An anonymous object has no constructor of its
                                // own, so its property initializers had nothing
                                // to run them — `object { val v = 4 }` left `v`
                                // undefined. Give it the one Kotlin gives it.
                                let mut inits = Vec::new();
                                for m in &mut body_members {
                                    if let ClassMember::Field {
                                        name: fname,
                                        init: field_init,
                                        ..
                                    } = m
                                    {
                                        if let Some(value) = field_init.take() {
                                            inits.push(Statement::new(StmtKind::Expr(
                                                Expression::new(ExprKind::Assign {
                                                    target: Box::new(Expression::new(
                                                        ExprKind::Member {
                                                            object: Box::new(Expression::new(
                                                                ExprKind::This,
                                                            )),
                                                            field: fname.clone(),
                                                            null_safe: false,
                                                        },
                                                    )),
                                                    value: Box::new(value),
                                                }),
                                            )));
                                        }
                                    }
                                }
                                members.extend(body_members);
                                if !inits.is_empty() {
                                    members.push(ClassMember::Constructor {
                                        name: None,
                                        params: Vec::new(),
                                        body: inits,
                                        base_args: None,
                                        initializer_target: ConstructorInitializerTarget::Base,
                                        visibility: Visibility::Public,
                                    });
                                }
                            }
                            _ => {}
                        }
                    }
                    // Kotlin's `object : I { … }` evaluates to an INSTANCE, not
                    // to the class. `ClassExpr` alone leaves the class object on
                    // the stack (that is JS's `class {}` expression), so every
                    // property read on it answered `undefined` and every method
                    // ran with the class as receiver.
                    Expression::new(ExprKind::New {
                        class: Box::new(Expression::new(ExprKind::ClassExpr {
                            name: None,
                            parent,
                            interfaces,
                            members,
                        })),
                        args: Vec::new(),
                    })
                }
                Rule::if_expr => {
                    let stmt = walk_if_stmt(inner).unwrap();
                    if let StmtKind::If {
                        cond,
                        then_body,
                        else_body,
                        ..
                    } = stmt.kind
                    {
                        let then_expr = then_body
                            .into_iter()
                            .last()
                            .and_then(|s| match s.kind {
                                StmtKind::Expr(e) => Some(e),
                                StmtKind::Return(Some(e)) => Some(e),
                                _ => None,
                            })
                            .unwrap_or_else(Expression::null);

                        let else_expr = else_body
                            .unwrap_or_default()
                            .into_iter()
                            .last()
                            .and_then(|s| match s.kind {
                                StmtKind::Expr(e) => Some(e),
                                StmtKind::Return(Some(e)) => Some(e),
                                _ => None,
                            })
                            .unwrap_or_else(Expression::null);

                        Expression::new(ExprKind::Ternary {
                            cond: Box::new(cond),
                            then: Box::new(then_expr),
                            else_: Box::new(else_expr),
                        })
                    } else {
                        Expression::null()
                    }
                }
                Rule::when_expr => {
                    let mut disc = None;
                    let mut entries = Vec::new();

                    for p in inner.into_inner() {
                        match p.as_rule() {
                            Rule::expr => disc = Some(walk_expr(p)),
                            Rule::when_entry => entries.push(p),
                            _ => {}
                        }
                    }

                    // See `walk_when_stmt`: `MatchArm::conditions` holds VALUES
                    // the subject may equal, so any arm that TESTS the subject
                    // (`is`, `in`, a range, a bare comparison) flips the whole
                    // expression to the subjectless shape.
                    let predicate_mode = disc.is_some()
                        && entries.iter().any(|entry| {
                            entry.clone().into_inner().any(|p| {
                                p.as_rule() == Rule::when_condition
                                    && when_condition_needs_predicate_expr(&p)
                            })
                        });
                    let subject = disc.clone().unwrap_or_else(|| Expression::bool(true));

                    let mut arms = Vec::new();

                    for entry in entries {
                        let mut entry_inner = entry.into_inner();
                        let mut is_else = false;
                        let mut cond_exprs = Vec::new();
                        let mut body_expr = Expression::null();

                        while let Some(p) = entry_inner.next() {
                            match p.as_rule() {
                                Rule::else_kw => is_else = true,
                                Rule::when_condition if predicate_mode => {
                                    if let Some(test) = when_condition_predicate(p, &subject) {
                                        cond_exprs.push(test);
                                    }
                                }
                                Rule::when_condition => {
                                    for csub in p.into_inner() {
                                        if csub.as_rule() == Rule::expr {
                                            cond_exprs.push(walk_expr(csub));
                                        }
                                    }
                                }
                                Rule::block => {
                                    let stmts = walk_block_statements(p);
                                    if let Some(last) = stmts.into_iter().last() {
                                        body_expr = match last.kind {
                                            StmtKind::Expr(e) => e,
                                            StmtKind::Return(Some(e)) => e,
                                            _ => Expression::null(),
                                        };
                                    }
                                }
                                Rule::statement => {
                                    if let Some(s) = walk_statement(p) {
                                        body_expr = match s.kind {
                                            StmtKind::Expr(e) => e,
                                            StmtKind::Return(Some(e)) => e,
                                            _ => Expression::null(),
                                        };
                                    }
                                }
                                _ => {}
                            }
                        }

                        let conditions = if is_else { None } else { Some(cond_exprs) };

                        arms.push(MatchArm {
                            conditions,
                            body: body_expr,
                        });
                    }

                    Expression::new(ExprKind::Match {
                        subject: Box::new(if predicate_mode {
                            Expression::bool(true)
                        } else {
                            subject
                        }),
                        arms,
                    })
                }
                Rule::expr => walk_expr(inner),
                _ => Expression::null(),
            }
        }
        _ => Expression::null(),
    }
}

fn walk_binary_chain(pair: Pair<Rule>, op: BinOp) -> Expression {
    let mut inner = pair.into_inner();
    let mut current = walk_expr(inner.next().unwrap());
    while let Some(_op_pair) = inner.next() {
        let next_expr = walk_expr(inner.next().unwrap());
        current = Expression::new(ExprKind::Binary {
            op,
            left: Box::new(current),
            right: Box::new(next_expr),
        });
    }
    current
}

/// The type a `type_ref` names, with Kotlin's nullability marker removed.
///
/// `String?` and `String` are ONE type as far as the shared machinery is
/// concerned — nullability is carried by `Param::is_nullable`, not by the
/// hint's spelling. Stripping it here is what lets `[builtin_types]` declare
/// each spelling once (`builtinslotplan.md` step 4a); leaving the `?` on would
/// make every declared spelling need a nullable twin, and `String?` would
/// resolve to no built-in at all.
fn type_hint_text(raw: &str) -> String {
    raw.trim().trim_end_matches('?').trim_end().to_string()
}

/// Whether a `type_ref`'s source text carries Kotlin's `?`.
fn type_ref_is_nullable(raw: &str) -> bool {
    raw.trim_end().ends_with('?')
}

/// Kotlin's infix spellings of the bitwise operators -> the shared `BinOp`.
///
/// `ushr` has no `BinOp` of its own; Kotlin's other shift is arithmetic, so
/// `shr` takes `Shr` and `ushr` stays a member call until an unsigned shift
/// exists to route it to.
fn infix_bitwise_op(op: &str) -> Option<BinOp> {
    match op {
        "and" => Some(BinOp::BitAnd),
        "or" => Some(BinOp::BitOr),
        "xor" => Some(BinOp::BitXor),
        "shl" => Some(BinOp::Shl),
        "shr" => Some(BinOp::Shr),
        _ => None,
    }
}

/// `0xFF` / `0b1010` / `1_000_000` / `12L` / `7u` -> the integer value.
///
/// The `_` grouping is a lexical convenience with no value, and the `u`/`L`
/// suffixes only pick the Kotlin static type; both are stripped before the
/// radix parse so one function covers every spelling.
fn parse_int_literal(raw: &str) -> i64 {
    let cleaned: String = raw.chars().filter(|c| *c != '_').collect();
    let body = cleaned.trim_end_matches(['L', 'l', 'u', 'U']);
    let (digits, radix) =
        if let Some(rest) = body.strip_prefix("0x").or_else(|| body.strip_prefix("0X")) {
            (rest, 16)
        } else if let Some(rest) = body.strip_prefix("0b").or_else(|| body.strip_prefix("0B")) {
            (rest, 2)
        } else {
            (body, 10)
        };
    i64::from_str_radix(digits, radix).unwrap_or(0)
}

fn walk_literal(pair: Pair<Rule>) -> Expression {
    let inner = pair.into_inner().next().unwrap();
    match inner.as_rule() {
        Rule::null_kw => Expression::null(),
        Rule::true_kw => Expression::bool(true),
        Rule::false_kw => Expression::bool(false),
        Rule::int_literal => Expression::int(parse_int_literal(inner.as_str())),
        Rule::float_literal => {
            let s: String = inner
                .as_str()
                .chars()
                .filter(|c| *c != '_')
                .collect::<String>()
                .trim_end_matches(['f', 'F'])
                .to_string();
            Expression::float(s.parse::<f64>().unwrap_or(0.0))
        }
        Rule::string_literal => walk_string_literal(inner),
        Rule::char_literal => {
            let s = inner.as_str();
            let content = &s[1..s.len().saturating_sub(1)];
            let decoded = match content {
                "\\n" => "\n".to_string(),
                "\\t" => "\t".to_string(),
                "\\r" => "\r".to_string(),
                "\\\"" => "\"".to_string(),
                "\\'" => "'".to_string(),
                "\\\\" => "\\".to_string(),
                "\\$" => "$".to_string(),
                "\\b" => "\x08".to_string(),
                "\\f" => "\x0C".to_string(),
                s if s.starts_with("\\u") && s.len() == 6 => {
                    if let Ok(code) = u32::from_str_radix(&s[2..], 16) {
                        if let Some(ch) = char::from_u32(code) {
                            ch.to_string()
                        } else {
                            s.to_string()
                        }
                    } else {
                        s.to_string()
                    }
                }
                s => s.to_string(),
            };
            Expression::string(&decoded)
        }
        _ => Expression::null(),
    }
}

fn walk_string_literal(pair: Pair<Rule>) -> Expression {
    let mut parts = Vec::new();
    collect_string_parts(pair, &mut parts);

    // `str_text` matches ONE character, so `"x="` arrives as two parts and a
    // plain literal was never a `Lit(Str)` at all — it was a `Binary` tree one
    // node per character. Everything downstream that asks "is this a string?"
    // (the `+` decision below, `[builtin_types]` classification, constant
    // folding) answered no for every literal longer than one char. Fold the
    // adjacent literal runs back into the single literal the source wrote.
    let mut folded: Vec<Expression> = Vec::new();
    for part in parts {
        match (&part.kind, folded.last_mut().map(|last| &mut last.kind)) {
            (ExprKind::Lit(Literal::Str(text)), Some(ExprKind::Lit(Literal::Str(acc)))) => {
                acc.push_str(text);
            }
            _ => folded.push(part),
        }
    }

    if folded.is_empty() {
        Expression::string("")
    } else if folded.len() == 1 {
        folded.remove(0)
    } else {
        // A template IS concatenation — `"a${x}b"` never means arithmetic, and
        // each interpolated part is already rendered by `__kt_tostring`.
        let mut iter = folded.into_iter();
        let mut acc = iter.next().unwrap();
        for p in iter {
            acc = Expression::new(ExprKind::Binary {
                op: BinOp::Concat,
                left: Box::new(acc),
                right: Box::new(p),
            });
        }
        acc
    }
}

/// One `$x` / `${expr}` inside a string template.
///
/// Kotlin templates call `toString()` on the part — they do NOT lean on `+`
/// coercion, and the two disagree: a `Boolean` concatenates as `1`, a `List` as
/// `1,2,3`, a `Map` as `[object]`. Routing the part through the same renderer
/// `println` uses is what makes `"v=$flag"` and `println(flag)` agree, which is
/// the whole point of having one renderer.
fn interpolated_part(expr: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("__kt_tostring")),
        args: vec![Argument::positional(expr)],
        optional: false,
    })
}

fn collect_string_parts(pair: Pair<Rule>, parts: &mut Vec<Expression>) {
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::raw_string_literal
            | Rule::plain_string_literal
            | Rule::string_content
            | Rule::raw_string_content => {
                collect_string_parts(child, parts);
            }
            Rule::str_text | Rule::raw_str_text => {
                parts.push(Expression::string(child.as_str()));
            }
            Rule::str_escaped => {
                let s = child.as_str();
                let decoded = match s {
                    "\\n" => "\n".to_string(),
                    "\\t" => "\t".to_string(),
                    "\\r" => "\r".to_string(),
                    "\\\"" => "\"".to_string(),
                    "\\'" => "'".to_string(),
                    "\\\\" => "\\".to_string(),
                    "\\$" => "$".to_string(),
                    "\\b" => "\x08".to_string(),
                    "\\f" => "\x0C".to_string(),
                    s if s.starts_with("\\u") && s.len() == 6 => {
                        if let Ok(code) = u32::from_str_radix(&s[2..], 16) {
                            if let Some(ch) = char::from_u32(code) {
                                ch.to_string()
                            } else {
                                s.to_string()
                            }
                        } else {
                            s.to_string()
                        }
                    }
                    _ => s.to_string(),
                };
                parts.push(Expression::string(&decoded));
            }
            Rule::str_interpolated_var => {
                if let Some(id_pair) = child.into_inner().next() {
                    let name = backing_field_substitution(id_pair.as_str());
                    let expr = inner_outer_read(&name)
                        .or_else(|| extension_receiver_read(&name))
                        .unwrap_or_else(|| Expression::ident(&name));
                    parts.push(interpolated_part(expr));
                }
            }
            Rule::str_interpolated_expr => {
                let raw = child.as_str();
                if raw.starts_with("${") && raw.ends_with('}') {
                    let inner_str = raw[2..raw.len() - 1].trim();
                    let unescaped = inner_str.replace("\\\"", "\"");
                    if let Ok(mut pairs) = KotlinParser::parse(Rule::expr, &unescaped) {
                        if let Some(epair) = pairs.next() {
                            parts.push(interpolated_part(walk_expr(epair)));
                        }
                    }
                }
            }
            _ => {}
        }
    }
}

/// `a to b` / `Pair(a, b)`.
///
/// `ExprKind::Tuple`, not `ExprKind::Array`: a Pair and a two-element List are
/// the same runtime array, and only the tuple TAG (`tuple_literals_tagged` in
/// the profile) tells them apart — which is what lets `println` render one as
/// `(3, x)` and the other as `[3, x]`.
fn create_pair_expr(a: Expression, b: Expression) -> Expression {
    Expression::new(ExprKind::Tuple(vec![a, b]))
}

/// `Triple(a, b, c)` — see [`create_pair_expr`].
fn create_triple_expr(a: Expression, b: Expression, c: Expression) -> Expression {
    Expression::new(ExprKind::Tuple(vec![a, b, c]))
}

/// `mapOf(…)` / a Kotlin map literal.
///
/// Plain data. This used to append a synthesised `toString` PROPERTY to the
/// object, which made the map render itself — and put `toString` into the
/// map's own `__keys`, so the map contained a member the program never put
/// there (`{a=1, b=2, toString=…}` once the renderer stopped hiding it).
/// Rendering is `emitter/tostring.rs`'s job; a map is just its entries.
fn create_map_expr(props: Vec<ObjectProperty>) -> Expression {
    Expression::new(ExprKind::Object(props))
}

/// `setOf(…)` / `mutableSetOf(…)`.
///
/// A Kotlin `Set` is a dict whose values are all `true` — the keys ARE the
/// elements, which is what gives `in` its O(1) answer. It carries
/// [`SET_MARKER`] because a `Set` and a `Map` are the same runtime shape and
/// render differently: `[1, 2, 3]` versus `{a=1}`.
fn create_kotlin_set_expr(elems: Vec<Expression>) -> Expression {
    let mut props = Vec::with_capacity(elems.len() + 1);
    props.push(ObjectProperty::KeyValue {
        key: Expression::string(SET_MARKER),
        value: Expression::bool(true),
    });
    for elem in elems {
        props.push(ObjectProperty::KeyValue {
            key: elem,
            value: Expression::bool(true),
        });
    }
    Expression::new(ExprKind::Object(props))
}
