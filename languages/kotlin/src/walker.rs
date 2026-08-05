use pest::Parser;
use pest::iterators::Pair;
use std::collections::{HashMap, HashSet};
use std::sync::atomic::{AtomicUsize, Ordering};
use vybe_ast::*;

use super::{KotlinParser, Rule};
use crate::emitter::tostring::SET_MARKER;

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
    /// Params of every SINGLY-declared top-level `fun`, for NAMED-argument
    /// reordering at call sites (overloads are excluded — named args on an
    /// overload group would need signature-aware selection first).
    static KOTLIN_TOP_FN_PARAMS: std::cell::RefCell<std::collections::HashMap<String, Vec<Param>>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
    /// Per top-level function NAME, in declaration order: each declaration's
    /// per-parameter RAW type text (generics survive here; `TypeHint`
    /// spellings are stripped to the bare head).
    static KOTLIN_FN_RAW_PARAM_TYPES: std::cell::RefCell<std::collections::HashMap<String, Vec<Vec<String>>>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
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
    /// Extension fns on a COMPANION (`fun Factory.Companion.from(…)`) —
    /// call sites spell them `Factory.from(…)`, keyed (class, name).
    static COMPANION_EXT_FNS: std::cell::RefCell<std::collections::HashSet<(String, String)>> =
        std::cell::RefCell::new(std::collections::HashSet::new());
    /// Declarations produced as a SIDE product of walking another one
    /// (the setter half of an extension property) — drained into the
    /// top level after collection.
    static PENDING_TOP_LEVEL_FNS: std::cell::RefCell<Vec<Statement>> =
        std::cell::RefCell::new(Vec::new());
    /// Extension properties WITH a setter — `holder.doubled = 10` calls
    /// `doubled__ext_set(holder, 10)`.
    static EXTENSION_PROPERTY_SETTERS: std::cell::RefCell<std::collections::HashSet<String>> =
        std::cell::RefCell::new(std::collections::HashSet::new());
    static EXTENSION_PROPERTIES: std::cell::RefCell<std::collections::HashSet<String>> =
        std::cell::RefCell::new(std::collections::HashSet::new());
    /// Class name -> the members it declares. An extension body reads the
    /// receiver's members unqualified (`val P.twice get() = n * 2`), so those
    /// names have to resolve to `this.<name>`.
    static CLASS_MEMBERS: std::cell::RefCell<std::collections::HashMap<String, std::collections::HashSet<String>>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
    static CLASS_PROPERTIES: std::cell::RefCell<std::collections::HashMap<String, std::collections::HashMap<String, bool>>> =
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
    static KOTLIN_SIMPLE_FUNCTIONS: std::cell::RefCell<std::collections::HashMap<String, (Vec<String>, Vec<Statement>)>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
    static KOTLIN_SEQUENCE_SOURCES: std::cell::RefCell<std::collections::HashMap<String, Expression>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
    static KOTLIN_STATIC_VALUES: std::cell::RefCell<std::collections::HashMap<String, Expression>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
    static KOTLIN_TUPLE_LOCALS: std::cell::RefCell<std::collections::HashSet<String>> =
        std::cell::RefCell::new(std::collections::HashSet::new());
    static KOTLIN_KEYED_COLLECTION_TYPES: std::cell::RefCell<std::collections::HashMap<String, String>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
    static KOTLIN_DATA_CLASS_PROPERTY_INDEX: std::cell::RefCell<std::collections::HashMap<String, std::collections::HashMap<String, usize>>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
    static KOTLIN_CLASS_PRIMARY_CTORS: std::cell::RefCell<std::collections::HashMap<String, Vec<Param>>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
    static KOTLIN_DELEGATED_COLLECTIONS: std::cell::RefCell<std::collections::HashMap<String, String>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
    static KOTLIN_NULLABLE_CTOR_CLASSES: std::cell::RefCell<std::collections::HashSet<String>> =
        std::cell::RefCell::new(std::collections::HashSet::new());
}

fn outer_this(depth: usize) -> Expression {
    let mut expr = Expression::new(ExprKind::This);
    for _ in 0..depth {
        expr = Expression::new(ExprKind::Member {
            object: Box::new(expr),
            field: "__kt_outer".to_string(),
            null_safe: false });
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
                        null_safe: false })
                })
        })
    })
}

fn qualified_inner_class(name: &str) -> Option<String> {
    INNER_CLASS_QUALIFIED.with(|m| m.borrow().get(name).cloned())
}

fn is_qualified_inner_class_path(path: &str) -> bool {
    INNER_CLASS_QUALIFIED.with(|m| m.borrow().values().any(|qualified| qualified == path))
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
            null_safe: false });
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
                    null_safe: false })
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
                // Every CALLABLE arity, not just the full one: `fun add(a:
                // Int, b: Int = 1)` answers 1-arg calls too, and the
                // spelling-rewrite guard (`is_user_member_name`) otherwise
                // let the collection `add` steal `acc.add(3)`.
                let defaults = inners
                    .iter()
                    .find(|p| p.as_rule() == Rule::parameter_list)
                    .map(|p| {
                        p.clone()
                            .into_inner()
                            .filter(|q| q.as_rule() == Rule::parameter)
                            .filter(|q| {
                                q.clone().into_inner().any(|r| r.as_rule() == Rule::expr)
                            })
                            .count()
                    })
                    .unwrap_or(0);
                for a in arity.saturating_sub(defaults)..arity {
                    out.insert((name.clone(), a));
                }
                out.insert((name, arity));
            }
        }
        if rule == Rule::function_decl {
            let inners: Vec<_> = pair.clone().into_inner().collect();
            if let Some(recv) = inners.iter().find(|p| p.as_rule() == Rule::receiver_prefix) {
                if let Some(id) = inners.iter().find(|p| p.as_rule() == Rule::identifier) {
                    EXTENSION_FUNCTIONS
                        .with(|set| set.borrow_mut().insert(id.as_str().to_string()));
                    // `Factory.Companion.` — the call site is the CLASS name.
                    let recv_text = recv.as_str().trim_end_matches('.');
                    if let Some(class) = recv_text.strip_suffix(".Companion") {
                        COMPANION_EXT_FNS.with(|set| {
                            set.borrow_mut()
                                .insert((class.to_string(), id.as_str().to_string()));
                        });
                    }
                }
            }
        }
        if rule == Rule::var_decl {
            let inners: Vec<_> = pair.clone().into_inner().collect();
            if inners.iter().any(|p| p.as_rule() == Rule::receiver_prefix) {
                if let Some(id) = inners.iter().find(|p| p.as_rule() == Rule::identifier) {
                    EXTENSION_PROPERTIES
                        .with(|set| set.borrow_mut().insert(id.as_str().to_string()));
                    let has_setter = inners.iter().any(|p| {
                        p.as_rule() == Rule::property_accessor
                            && p.clone().into_inner().any(|q| q.as_rule() == Rule::set_kw)
                    });
                    if has_setter {
                        EXTENSION_PROPERTY_SETTERS
                            .with(|set| set.borrow_mut().insert(id.as_str().to_string()));
                    }
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
                    let readonly = rule == Rule::class_parameter
                        && inners.iter().any(|p| p.as_rule() == Rule::val_kw)
                        || rule == Rule::var_decl
                            && inners.iter().any(|p| p.as_rule() == Rule::val_kw);
                    record_member(id.as_str());
                    if let Some(owner) = owner_ref {
                        CLASS_PROPERTIES.with(|m| {
                            m.borrow_mut()
                                .entry(owner.to_string())
                                .or_default()
                                .insert(id.as_str().to_string(), readonly);
                        });
                    }
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
    COMPANION_EXT_FNS.with(|set| set.borrow_mut().clear());
    EXTENSION_PROPERTY_SETTERS.with(|set| set.borrow_mut().clear());
    CLASS_MEMBERS.with(|m| m.borrow_mut().clear());
    CLASS_PROPERTIES.with(|m| m.borrow_mut().clear());
    KOTLIN_DELEGATED_COLLECTIONS.with(|m| m.borrow_mut().clear());
    KOTLIN_NULLABLE_CTOR_CLASSES.with(|set| set.borrow_mut().clear());
    USER_METHOD_OVERLOADS.with(|map| map.borrow_mut().clear());
    KOTLIN_FN_RAW_PARAM_TYPES.with(|map| map.borrow_mut().clear());
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

fn has_nullable_constructor_param(name: &str) -> bool {
    KOTLIN_NULLABLE_CTOR_CLASSES.with(|set| set.borrow().contains(name))
}

fn kotlin_user_constructor_call(name: &str, args: &[Argument]) -> Option<Expression> {
    if !is_user_class_name(name)
        || !has_nullable_constructor_param(name)
        || args.iter().any(|arg| arg.name.is_some())
    {
        return None;
    }
    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(&format!("{name}$arity{}", args.len()))),
        args: args.to_vec(),
        optional: false }))
}

fn kotlin_normalized_constructor_args(name: &str, args: &[Argument]) -> Vec<Argument> {
    if args.is_empty()
        || args.iter().all(|arg| arg.name.is_none())
        || args.iter().any(|arg| arg.spread)
    {
        return args.to_vec();
    }

    let Some(params) = KOTLIN_CLASS_PRIMARY_CTORS.with(|ctors| ctors.borrow().get(name).cloned())
    else {
        return args.to_vec();
    };

    let mut positional = args.iter().filter(|arg| arg.name.is_none());
    let mut out = Vec::new();
    for param in params {
        if let Some(named) = args
            .iter()
            .find(|arg| arg.name.as_deref() == Some(param.name.as_str()))
        {
            out.push(Argument::positional(named.value.clone()));
        } else if let Some(arg) = positional.next() {
            out.push(Argument::positional(arg.value.clone()));
        } else if let Some(default) = param.default {
            out.push(Argument::positional(default));
        } else {
            out.push(Argument::positional(Expression::null()));
        }
    }
    out.extend(positional.map(|arg| Argument::positional(arg.value.clone())));
    out
}

/// NAMED arguments on a top-level `fun` call — reorder against the declared
/// parameter list, filling skipped middles with their defaults (the shared
/// call path only fills TRAILING defaults positionally).
fn kotlin_normalized_fn_args(name: &str, args: &[Argument]) -> Option<Vec<Argument>> {
    if args.iter().all(|arg| arg.name.is_none()) || args.iter().any(|arg| arg.spread) {
        return None;
    }
    let params = KOTLIN_TOP_FN_PARAMS.with(|map| map.borrow().get(name).cloned())?;
    let mut positional = args.iter().filter(|arg| arg.name.is_none());
    let mut out = Vec::new();
    for param in &params {
        if let Some(named) = args
            .iter()
            .find(|arg| arg.name.as_deref() == Some(param.name.as_str()))
        {
            out.push(Argument::positional(named.value.clone()));
        } else if let Some(arg) = positional.next() {
            out.push(Argument::positional(arg.value.clone()));
        } else if let Some(default) = &param.default {
            out.push(Argument::positional(default.clone()));
        } else {
            out.push(Argument::positional(Expression::null()));
        }
    }
    out.extend(positional.map(|arg| Argument::positional(arg.value.clone())));
    Some(out)
}

/// Whether a class in this source declares a method with this name and arity.
fn is_user_member_name(name: &str, arity: usize) -> bool {
    USER_MEMBER_NAMES.with(|set| set.borrow().contains(&(name.to_string(), arity)))
}

fn overloaded_storage_name(name: &str, arity: usize) -> Option<String> {
    overloaded_storage_name_for_args(name, arity, &[])
}

/// Pick the method-overload signature for a call. One signature per arity is
/// unambiguous; SAME-arity overloads (`eval(Int)` vs `eval(String)`) resolve
/// by classifying the literal/lambda arguments.
fn overloaded_storage_name_for_args(
    name: &str,
    arity: usize,
    args: &[Argument],
) -> Option<String> {
    USER_METHOD_OVERLOADS.with(|map| {
        let map = map.borrow();
        let signatures = map.get(name)?.get(&arity)?;
        let param_types = if signatures.len() == 1 {
            &signatures[0]
        } else {
            let arg_tys: Vec<Option<OvlTy>> =
                args.iter().map(|a| ovl_arg_ty(&a.value)).collect();
            if args.len() != arity || arg_tys.iter().all(Option::is_none) {
                return None;
            }
            signatures.iter().find(|sig| {
                sig.iter().zip(arg_tys.iter()).all(|(p, a)| match a {
                    Some(t) => ovl_param_ty(Some(p.as_str())) == *t,
                    None => true })
            })?
        };
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
                                        body: Box::new(stmt) }));
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

    PENDING_TOP_LEVEL_FNS.with(|pending| body.extend(pending.borrow_mut().drain(..)));

    let aliases = kotlin_import_aliases(&imports);
    rewrite_import_aliases_in_stmts(&mut body, &aliases);
    {
        let mut counts: HashMap<String, usize> = HashMap::new();
        for stmt in &body {
            if let StmtKind::FunctionDecl { name, .. } = &stmt.kind {
                *counts.entry(name.clone()).or_default() += 1;
            }
        }
        KOTLIN_TOP_FN_PARAMS.with(|map| {
            let mut map = map.borrow_mut();
            map.clear();
            for stmt in &body {
                if let StmtKind::FunctionDecl { name, params, .. } = &stmt.kind {
                    if counts.get(name.as_str()) == Some(&1) {
                        map.insert(name.clone(), params.clone());
                    }
                }
            }
        });
        // Resolve the `::fn` markers now that the table exists: a user
        // function's ref IS the function value; anything else keeps the
        // one-argument wrapper lambda.
        for stmt in body.iter_mut() {
            stmt.walk_exprs_mut(&mut |e| {
                if let ExprKind::Call { callee, args, .. } = &e.kind
                    && matches!(&callee.kind, ExprKind::Ident(n) if n == "__kt_callable_ref")
                    && let Some(Argument { value, .. }) = args.first()
                    && let ExprKind::Lit(Literal::Str(target)) = &value.kind
                {
                    let target = target.to_string();
                    *e = if KOTLIN_TOP_FN_PARAMS
                        .with(|map| map.borrow().contains_key(target.as_str()))
                    {
                        Expression::ident(&target)
                    } else {
                        callable_ref_lambda(Expression::ident(&target))
                    };
                }
            });
        }
    }
    kotlin_lower_elvis_returns(&mut body);
    kotlin_unlift_try_inits(&mut body);
    {
        // Seed with builtin FREE-call spellings the emitter intercepts by
        // name: a local `val toString = { … }` must shadow the builtin (its
        // call sites rename with it), or `toString(3)` keeps converting.
        let mut top = HashSet::from(["toString".to_string()]);
        let _ = kotlin_rename_shadowed_decls(&mut body, &mut top, &mut 0);
    }
    kotlin_hoist_local_fns(&mut body);
    kotlin_resolve_function_overloads(&mut body);
    collect_kotlin_simple_functions(&body);
    KOTLIN_SEQUENCE_SOURCES.with(|map| map.borrow_mut().clear());
    KOTLIN_STATIC_VALUES.with(|map| map.borrow_mut().clear());
    KOTLIN_TUPLE_LOCALS.with(|set| set.borrow_mut().clear());
    KOTLIN_KEYED_COLLECTION_TYPES.with(|map| map.borrow_mut().clear());
    normalize_kotlin_operator_calls(&mut body);
    KOTLIN_STATIC_VALUES.with(|map| map.borrow_mut().clear());
    KOTLIN_TUPLE_LOCALS.with(|set| set.borrow_mut().clear());
    KOTLIN_SEQUENCE_SOURCES.with(|map| map.borrow_mut().clear());
    KOTLIN_KEYED_COLLECTION_TYPES.with(|map| map.borrow_mut().clear());
    KOTLIN_DATA_CLASS_PROPERTY_INDEX.with(|map| map.borrow_mut().clear());
    KOTLIN_CLASS_PRIMARY_CTORS.with(|map| map.borrow_mut().clear());
    KOTLIN_DELEGATED_COLLECTIONS.with(|map| map.borrow_mut().clear());
    KOTLIN_NULLABLE_CTOR_CLASSES.with(|set| set.borrow_mut().clear());
    KOTLIN_SIMPLE_FUNCTIONS.with(|map| map.borrow_mut().clear());

    // An `enum class`'s constants are built by its `__static_init_block__`
    // (platforms/jvm `lang_enum`), and a static initializer only runs because
    // something calls it. Before the package wrapper below, so the calls sit
    // beside the declarations rather than outside the namespace.
    vybe_platform_jvm::lang_enum::inject_static_init_calls(&mut body);

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
                    optional: false },
            ))));
        }
    }

    Ok(Module {
        name: "main".to_string(),
        language: Lang::Kotlin,
        body,
        imports })
}

fn collect_kotlin_simple_functions(stmts: &[Statement]) {
    KOTLIN_SIMPLE_FUNCTIONS.with(|map| {
        let mut map = map.borrow_mut();
        map.clear();
        for stmt in stmts {
            if let StmtKind::FunctionDecl {
                name, params, body, ..
            } = &stmt.kind
            {
                let param_names = params.iter().map(|param| param.name.clone()).collect();
                map.insert(name.clone(), (param_names, body.clone()));
            }
        }
    });
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
            null_safe: false })
    })
}

fn dotted_expr_path(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => Some(name.clone()),
        ExprKind::Member { object, field, .. } => {
            let mut path = dotted_expr_path(object)?;
            path.push('.');
            path.push_str(field);
            Some(path)
        }
        _ => None }
}

/// A `val`/`var` that REUSES a bound name is a NEW binding in Kotlin — both
/// the same-block redeclare (`var v = 1 … val v = 9`) and a nested block's
/// shadow (`try { val v = "inner" }`). The flat function scope made the new
/// declaration overwrite the old one, so the outer name never came back.
/// Renames the shadowing declaration and every LATER reference in its scope;
/// lambda parameters that rebind the name keep their own bodies untouched.
fn kotlin_rename_shadowed_decls(
    stmts: &mut Vec<Statement>,
    bound: &mut HashSet<String>,
    counter: &mut usize,
) -> Vec<(String, String)> {
    let mut leaked: Vec<(String, String)> = Vec::new();
    let mut i = 0;
    while i < stmts.len() {
        // Recurse into nested blocks with the CURRENT bindings.
        match &mut stmts[i].kind {
            StmtKind::FunctionDecl { name, params, body, .. } => {
                bound.insert(name.clone());
                let mut inner = bound.clone();
                for p in params.iter() {
                    inner.insert(p.name.clone());
                }
                kotlin_rename_shadowed_decls(body, &mut inner, counter);
            }
            StmtKind::If { then_body, elifs, else_body, .. } => {
                kotlin_rename_shadowed_decls(then_body, &mut bound.clone(), counter);
                for (_, b) in elifs {
                    kotlin_rename_shadowed_decls(b, &mut bound.clone(), counter);
                }
                if let Some(b) = else_body {
                    kotlin_rename_shadowed_decls(b, &mut bound.clone(), counter);
                }
            }
            StmtKind::While { body, .. } => {
                kotlin_rename_shadowed_decls(body, &mut bound.clone(), counter);
            }
            // A Block here is the walker's own TRANSPARENT lowering
            // (destructuring declarations) — its bindings live in the
            // ENCLOSING scope, so its renames continue past it.
            StmtKind::Block(body) => {
                let inner_leaks = kotlin_rename_shadowed_decls(body, bound, counter);
                for (old, new) in inner_leaks {
                    for s in stmts[i + 1..].iter_mut() {
                        kotlin_rename_ident_in_stmt(s, &old, &new);
                    }
                    bound.insert(new.clone());
                    leaked.push((old, new));
                }
            }
            StmtKind::ForIn { var, body, .. } => {
                let mut inner = bound.clone();
                inner.insert(var.clone());
                kotlin_rename_shadowed_decls(body, &mut inner, counter);
            }
            StmtKind::Try { body, catches, else_body, finally } => {
                kotlin_rename_shadowed_decls(body, &mut bound.clone(), counter);
                for c in catches {
                    kotlin_rename_shadowed_decls(&mut c.body, &mut bound.clone(), counter);
                }
                if let Some(b) = else_body {
                    kotlin_rename_shadowed_decls(b, &mut bound.clone(), counter);
                }
                if let Some(b) = finally {
                    kotlin_rename_shadowed_decls(b, &mut bound.clone(), counter);
                }
            }
            StmtKind::VarDecl { declarations, .. } => {
                let mut renames: Vec<(String, String)> = Vec::new();
                for d in declarations.iter_mut() {
                    if let BindingPattern::Ident(name) = &mut d.pattern {
                        if bound.contains(name.as_str()) {
                            let renamed = format!("{name}__shadow{counter}");
                            *counter += 1;
                            renames.push((name.clone(), renamed.clone()));
                            *name = renamed;
                        } else {
                            bound.insert(name.clone());
                        }
                    }
                }
                for (old, new) in renames {
                    for s in stmts[i + 1..].iter_mut() {
                        kotlin_rename_ident_in_stmt(s, &old, &new);
                    }
                    bound.insert(new.clone());
                    leaked.push((old, new));
                }
            }
            _ => {}
        }
        // Lambda bodies are real nested scopes reached only through
        // EXPRESSIONS (`run { … }`, callbacks) — without this, a
        // `val value = value * 10` inside `run {}` resolved its own INIT to
        // the not-yet-set lambda local (NaN) instead of the outer binding.
        // The deep visitor re-reaches lambdas the recursion already handled;
        // that is harmless — a renamed declaration no longer collides.
        {
            let bound_now = bound.clone();
            stmts[i].walk_exprs_mut(&mut |e| {
                if let ExprKind::Lambda { params, body, .. } = &mut e.kind {
                    if let LambdaBody::Block(b) = body {
                        let mut inner = bound_now.clone();
                        for p in params.iter() {
                            inner.insert(p.name.clone());
                        }
                        kotlin_rename_shadowed_decls(b, &mut inner, counter);
                    }
                }
            });
        }
        i += 1;
    }
    leaked
}

/// Rename inside a nested BLOCK, stopping when the block redeclares the
/// name itself (only that declaration's INIT still sees the outer binding).
fn kotlin_rename_ident_in_block(stmts: &mut [Statement], old: &str, new: &str) {
    for s in stmts {
        if let StmtKind::VarDecl { declarations, .. } = &mut s.kind {
            let mut rebinds = false;
            for d in declarations.iter_mut() {
                if let Some(init) = &mut d.init {
                    kotlin_rename_ident_in_expr(init, old, new);
                }
                if matches!(&d.pattern, BindingPattern::Ident(n) if n == old) {
                    rebinds = true;
                }
            }
            if rebinds {
                return;
            }
            continue;
        }
        kotlin_rename_ident_in_stmt(s, old, new);
    }
}

/// Rename identifier READS/WRITES of `old`, stopping at scopes that REBIND
/// the name (lambda params, nested `fun` params, a fresh `val old`).
fn kotlin_rename_ident_in_stmt(stmt: &mut Statement, old: &str, new: &str) {
    match &mut stmt.kind {
        StmtKind::Expr(e) | StmtKind::Return(Some(e)) | StmtKind::Throw { expr: Some(e), .. } => {
            kotlin_rename_ident_in_expr(e, old, new)
        }
        StmtKind::VarDecl { declarations, .. } => {
            for d in declarations {
                if let Some(init) = &mut d.init {
                    kotlin_rename_ident_in_expr(init, old, new);
                }
                // A LATER redeclaration of the same original name rebinds it —
                // stop renaming from here on is handled by the caller's
                // sequential order (its own rename pass runs afterwards).
            }
        }
        StmtKind::FunctionDecl { params, body, .. } => {
            if !params.iter().any(|p| p.name == old) {
                for s in body {
                    kotlin_rename_ident_in_stmt(s, old, new);
                }
            }
        }
        StmtKind::If { cond, then_body, elifs, else_body, .. } => {
            kotlin_rename_ident_in_expr(cond, old, new);
            kotlin_rename_ident_in_block(then_body, old, new);
            for (c, b) in elifs {
                kotlin_rename_ident_in_expr(c, old, new);
                kotlin_rename_ident_in_block(b, old, new);
            }
            if let Some(b) = else_body {
                kotlin_rename_ident_in_block(b, old, new);
            }
        }
        StmtKind::While { cond, body, .. } => {
            kotlin_rename_ident_in_expr(cond, old, new);
            kotlin_rename_ident_in_block(body, old, new);
        }
        StmtKind::ForIn { var, iter, body, .. } => {
            kotlin_rename_ident_in_expr(iter, old, new);
            if var != old {
                kotlin_rename_ident_in_block(body, old, new);
            }
        }
        StmtKind::Try { body, catches, else_body, finally } => {
            kotlin_rename_ident_in_block(body, old, new);
            for c in catches {
                kotlin_rename_ident_in_block(&mut c.body, old, new);
            }
            if let Some(b) = else_body {
                kotlin_rename_ident_in_block(b, old, new);
            }
            if let Some(b) = finally {
                kotlin_rename_ident_in_block(b, old, new);
            }
        }
        StmtKind::Block(body) => {
            kotlin_rename_ident_in_block(body, old, new);
        }
        _ => {}
    }
}

fn kotlin_rename_ident_in_expr(expr: &mut Expression, old: &str, new: &str) {
    match &mut expr.kind {
        ExprKind::Ident(n) if n == old => {
            *expr = Expression::ident(new);
        }
        ExprKind::Lambda { params, body, .. } => {
            if params.iter().any(|p| p.name == old) {
                return; // the parameter rebinds the name
            }
            match body {
                LambdaBody::Expr(e) => kotlin_rename_ident_in_expr(e, old, new),
                LambdaBody::Block(stmts) => {
                    kotlin_rename_ident_in_block(stmts, old, new);
                }
            }
        }
        ExprKind::Call { callee, args, .. } => {
            kotlin_rename_ident_in_expr(callee, old, new);
            for a in args {
                kotlin_rename_ident_in_expr(&mut a.value, old, new);
            }
        }
        ExprKind::Binary { left, right, .. } => {
            kotlin_rename_ident_in_expr(left, old, new);
            kotlin_rename_ident_in_expr(right, old, new);
        }
        ExprKind::Assign { target, value } => {
            kotlin_rename_ident_in_expr(target, old, new);
            kotlin_rename_ident_in_expr(value, old, new);
        }
        ExprKind::Member { object, .. } => {
            kotlin_rename_ident_in_expr(object, old, new);
        }
        ExprKind::Index { object, index, .. } => {
            kotlin_rename_ident_in_expr(object, old, new);
            kotlin_rename_ident_in_expr(index, old, new);
        }
        ExprKind::Unary { expr: inner, .. }
        | ExprKind::Cast { expr: inner, .. }
        | ExprKind::Await(inner)
        | ExprKind::Yield(Some(inner))
        | ExprKind::Delete(inner) => kotlin_rename_ident_in_expr(inner, old, new),
        ExprKind::Ternary { cond, then, else_ } => {
            kotlin_rename_ident_in_expr(cond, old, new);
            kotlin_rename_ident_in_expr(then, old, new);
            kotlin_rename_ident_in_expr(else_, old, new);
        }
        ExprKind::NullCoalesce { left, right } => {
            kotlin_rename_ident_in_expr(left, old, new);
            kotlin_rename_ident_in_expr(right, old, new);
        }
        ExprKind::Range { start, end, .. } => {
            kotlin_rename_ident_in_expr(start, old, new);
            kotlin_rename_ident_in_expr(end, old, new);
        }
        ExprKind::Array(items) => {
            for it in items {
                kotlin_rename_ident_in_expr(&mut it.value, old, new);
            }
        }
        ExprKind::Tuple(items) | ExprKind::Sequence(items) => {
            for it in items {
                kotlin_rename_ident_in_expr(it, old, new);
            }
        }
        ExprKind::Match { subject, arms } => {
            kotlin_rename_ident_in_expr(subject, old, new);
            for arm in arms {
                if let Some(conds) = &mut arm.conditions {
                    for c in conds {
                        kotlin_rename_ident_in_expr(c, old, new);
                    }
                }
                kotlin_rename_ident_in_expr(&mut arm.body, old, new);
            }
        }
        _ => {}
    }
}

/// `v ?: return x` — the walker marks the RHS as `__kt_elvis_return(x)`;
/// this pass lowers each occurrence to a GUARD statement before its host:
/// `if (v == null) return x` and the coalesce collapses to `v`.
fn kotlin_lower_elvis_returns(stmts: &mut Vec<Statement>) {
    let mut i = 0;
    while i < stmts.len() {
        // Recurse into nested blocks first.
        match &mut stmts[i].kind {
            StmtKind::FunctionDecl { body, .. } | StmtKind::While { body, .. }
            | StmtKind::ForIn { body, .. } | StmtKind::Block(body) => {
                kotlin_lower_elvis_returns(body);
            }
            StmtKind::If { then_body, elifs, else_body, .. } => {
                kotlin_lower_elvis_returns(then_body);
                for (_, b) in elifs {
                    kotlin_lower_elvis_returns(b);
                }
                if let Some(b) = else_body {
                    kotlin_lower_elvis_returns(b);
                }
            }
            _ => {}
        }
        let mut guards: Vec<(Expression, Option<Expression>)> = Vec::new();
        stmts[i].walk_exprs_mut(&mut |e| {
            kotlin_extract_elvis_return(e, &mut guards);
        });
        if !guards.is_empty() {
            for (left, payload) in guards.into_iter().rev() {
                let guard = Statement::new(StmtKind::If {
                    cond: Expression::new(ExprKind::Binary {
                        op: BinOp::Eq,
                        left: Box::new(left),
                        right: Box::new(Expression::null()) }),
                    then_body: vec![Statement::new(StmtKind::Return(payload))],
                    elifs: vec![],
                    else_body: None });
                stmts.insert(i, guard);
                i += 1;
            }
        }
        i += 1;
    }
}

fn kotlin_extract_elvis_return(
    expr: &mut Expression,
    guards: &mut Vec<(Expression, Option<Expression>)>,
) {
    if let ExprKind::NullCoalesce { left, right } = &expr.kind
        && let ExprKind::Call { callee, args, .. } = &right.kind
        && matches!(&callee.kind, ExprKind::Ident(n) if n == "__kt_elvis_return")
    {
        let payload = args.first().map(|a| a.value.clone());
        let l = (**left).clone();
        guards.push((l.clone(), payload));
        *expr = l;
    }
}

/// An `if` STATEMENT reshaped as its expression value — recursively, so
/// `if a X else if b Y else Z` chains keep every branch (the old one-level
/// conversion answered null for the else-if arm).
fn kotlin_if_stmt_to_ternary(stmt: Statement) -> Expression {
    let StmtKind::If { cond, then_body, else_body, .. } = stmt.kind else {
        return Expression::null();
    };
    fn branch_value(mut stmts: Vec<Statement>) -> Expression {
        if stmts.len() > 1 {
            // A multi-statement branch keeps its leading statements (local
            // fns, side effects) by becoming a zero-arg IIFE — the old
            // last-expression pick silently dropped them.
            kotlin_return_wrap_last(&mut stmts);
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Lambda {
                    params: Vec::new(),
                    body: LambdaBody::Block(stmts),
                    captures: Vec::new(),
                    is_async: false })),
                args: Vec::new(),
                optional: false });
        }
        stmts
            .into_iter()
            .last()
            .map(|s| match s.kind {
                StmtKind::Expr(e) => e,
                StmtKind::Return(Some(e)) => e,
                kind @ StmtKind::If { .. } => {
                    kotlin_if_stmt_to_ternary(Statement::new(kind))
                }
                // `if (bad) throw X else v` as a VALUE: the throw branch
                // rides the throw-expression helper (it answered null and
                // the exception vanished).
                StmtKind::Throw { expr: Some(e), .. } => Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__kt_throw")),
                    args: vec![Argument::positional(e)],
                    optional: false }),
                _ => Expression::null() })
            .unwrap_or_else(Expression::null)
    }
    Expression::new(ExprKind::Ternary {
        cond: Box::new(cond),
        then: Box::new(branch_value(then_body)),
        else_: Box::new(branch_value(else_body.unwrap_or_default())) })
}

/// Hoist CAPTURE-FREE local `fun`s to the top level under unique names.
///
/// A nested `fun fact(n) = … fact(n - 1)` compiled as a local closure whose
/// self-reference was null; hoisted, it is an ordinary recursive function.
/// Locals that read enclosing bindings stay put (their capture semantics are
/// the compiler's). Call sites rewrite only when the hoisted signature
/// accepts the arguments — a shadowing local with a narrower type must not
/// swallow calls meant for a same-name global.
fn kotlin_hoist_local_fns(body: &mut Vec<Statement>) {
    let mut hoisted: Vec<Statement> = Vec::new();
    let mut counter = 0usize;
    let top_fn_names: HashSet<String> = body
        .iter()
        .filter_map(|s| match &s.kind {
            StmtKind::FunctionDecl { name, .. } => Some(name.clone()),
            _ => None })
        .collect();
    for stmt in body.iter_mut() {
        if let StmtKind::FunctionDecl { params, body: fn_body, .. } = &mut stmt.kind {
            let mut enclosing: HashSet<String> =
                params.iter().map(|p| p.name.clone()).collect();
            kotlin_hoist_in_block(
                fn_body,
                &mut enclosing,
                &top_fn_names,
                &mut hoisted,
                &mut counter,
            );
        }
    }
    body.extend(hoisted);
}

fn kotlin_collect_ident_reads(stmts: &[Statement], out: &mut HashSet<String>) {
    for s in stmts {
        let mut s2 = s.clone();
        s2.walk_exprs_mut(&mut |e| {
            collect_idents_in_expr(e, out);
        });
        if let StmtKind::FunctionDecl { body, .. } = &s.kind {
            kotlin_collect_ident_reads(body, out);
        }
    }
}

fn collect_idents_in_expr(expr: &Expression, out: &mut HashSet<String>) {
    if let ExprKind::Ident(n) = &expr.kind {
        out.insert(n.clone());
    }
}

fn kotlin_collect_assigned_names(stmts: &[Statement], out: &mut HashSet<String>) {
    for s in stmts {
        let mut s2 = s.clone();
        s2.walk_exprs_mut(&mut |e| {
            if let ExprKind::Assign { target, .. } = &e.kind
                && let ExprKind::Ident(n) = &target.kind
            {
                out.insert(n.clone());
            }
            if let ExprKind::Unary { op, expr: inner } = &e.kind
                && matches!(op, UnaryOp::PostInc | UnaryOp::PostDec)
                && let ExprKind::Ident(n) = &inner.kind
            {
                out.insert(n.clone());
            }
        });
        if let StmtKind::FunctionDecl { body, .. } = &s.kind {
            kotlin_collect_assigned_names(body, out);
        }
    }
}

fn kotlin_local_binding_names(stmts: &[Statement], out: &mut HashSet<String>) {
    for s in stmts {
        match &s.kind {
            StmtKind::VarDecl { declarations, .. } => {
                for d in declarations {
                    if let BindingPattern::Ident(n) = &d.pattern {
                        out.insert(n.clone());
                    }
                }
            }
            StmtKind::If { then_body, elifs, else_body, .. } => {
                kotlin_local_binding_names(then_body, out);
                for (_, b) in elifs {
                    kotlin_local_binding_names(b, out);
                }
                if let Some(b) = else_body {
                    kotlin_local_binding_names(b, out);
                }
            }
            StmtKind::While { body, .. } | StmtKind::ForIn { body, .. } => {
                kotlin_local_binding_names(body, out);
            }
            StmtKind::FunctionDecl { name, params, body, .. } => {
                out.insert(name.clone());
                for p in params {
                    out.insert(p.name.clone());
                }
                kotlin_local_binding_names(body, out);
            }
            _ => {}
        }
    }
}

fn kotlin_hoist_in_block(
    stmts: &mut Vec<Statement>,
    enclosing: &mut HashSet<String>,
    top_fn_names: &HashSet<String>,
    hoisted: &mut Vec<Statement>,
    counter: &mut usize,
) {
    // Bindings this block introduces become "enclosing" for nested blocks.
    let mut i = 0;
    while i < stmts.len() {
        let is_local_fn = matches!(stmts[i].kind, StmtKind::FunctionDecl { .. });
        if is_local_fn {
            // Free identifiers naming enclosing bindings become THREADED
            // captures — hoistable as long as the fn never ASSIGNS them
            // (read-only captures pass as leading parameters).
            let (fn_name, captured, writes_captures) = {
                let StmtKind::FunctionDecl { name, params, body, .. } = &stmts[i].kind else {
                    unreachable!()
                };
                let mut reads = HashSet::new();
                kotlin_collect_ident_reads(body, &mut reads);
                let mut own: HashSet<String> =
                    params.iter().map(|p| p.name.clone()).collect();
                own.insert(name.clone());
                kotlin_local_binding_names(body, &mut own);
                let mut captured: Vec<String> = reads
                    .iter()
                    .filter(|r| !own.contains(*r) && enclosing.contains(*r))
                    .cloned()
                    .collect();
                captured.sort();
                let mut assigned = HashSet::new();
                kotlin_collect_assigned_names(body, &mut assigned);
                let writes = captured.iter().any(|c| assigned.contains(c));
                (name.clone(), captured, writes)
            };
            if !writes_captures {
                let other_candidates = top_fn_names.contains(&fn_name)
                    || stmts[i + 1..].iter().any(|s| {
                        matches!(&s.kind, StmtKind::FunctionDecl { name, .. } if *name == fn_name)
                    });
                let mut decl = stmts.remove(i);
                let new_name = format!("{fn_name}__loc{counter}");
                *counter += 1;
                let sig;
                if let StmtKind::FunctionDecl { name, params, .. } = &mut decl.kind {
                    *name = new_name.clone();
                    for (k, c) in captured.iter().enumerate() {
                        params.insert(
                            k,
                            Param {
                                name: c.clone(),
                                type_hint: None,
                                default: None,
                                pass_by: PassBy::Value,
                                is_rest: false,
                                is_kwargs: false,
                                is_optional: false,
                                is_nullable: false },
                        );
                    }
                    sig = ovl_sig_of_params(&new_name, params);
                } else {
                    unreachable!()
                }
                // Self-recursion inside the hoisted body follows the rename
                // (and threads its own captures).
                decl.walk_exprs_mut(&mut |e| {
                    kotlin_rewrite_hoisted_ref(e, &fn_name, &sig, other_candidates, &captured);
                });
                // Rewrite the REST of this block: calls the local accepts,
                // and bare value references.
                for s in stmts[i..].iter_mut() {
                    s.walk_exprs_mut(&mut |e| {
                        kotlin_rewrite_hoisted_ref(e, &fn_name, &sig, other_candidates, &captured);
                    });
                }
                hoisted.push(decl);
                continue; // same index now holds the next statement
            }
        }
        match &mut stmts[i].kind {
            StmtKind::VarDecl { declarations, .. } => {
                for d in declarations {
                    if let BindingPattern::Ident(n) = &d.pattern {
                        enclosing.insert(n.clone());
                    }
                }
            }
            StmtKind::If { then_body, elifs, else_body, .. } => {
                kotlin_hoist_in_block(then_body, enclosing, top_fn_names, hoisted, counter);
                for (_, b) in elifs {
                    kotlin_hoist_in_block(b, enclosing, top_fn_names, hoisted, counter);
                }
                if let Some(b) = else_body {
                    kotlin_hoist_in_block(b, enclosing, top_fn_names, hoisted, counter);
                }
            }
            StmtKind::While { body, .. } | StmtKind::ForIn { body, .. } => {
                kotlin_hoist_in_block(body, enclosing, top_fn_names, hoisted, counter);
            }
            StmtKind::FunctionDecl { params, body, .. } => {
                // A CAPTURING local fn: hoist inside it too, with its params
                // joining the enclosing set.
                let mut inner = enclosing.clone();
                for p in params.iter() {
                    inner.insert(p.name.clone());
                }
                kotlin_hoist_in_block(body, &mut inner, top_fn_names, hoisted, counter);
            }
            _ => {}
        }
        i += 1;
    }
}

fn kotlin_rewrite_hoisted_ref(
    expr: &mut Expression,
    old: &str,
    sig: &OvlSig,
    other_candidates: bool,
    captured: &[String],
) {
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            if let ExprKind::Ident(n) = &callee.kind
                && n == old
            {
                let argc = args.len() + captured.len();
                let arg_tys: Vec<Option<OvlTy>> =
                    args.iter().map(|a| ovl_arg_ty(&a.value)).collect();
                let accepts = sig.min_args <= argc
                    && argc <= sig.max_args
                    && ovl_sig_accepts_offset(sig, args, &arg_tys, captured.len());
                // With no other same-name candidate (top-level fn or a later
                // local overload), the local is the only resolution.
                if accepts || !other_candidates {
                    **callee = Expression::ident(&sig.mangled);
                    // Captured enclosing bindings thread through as leading
                    // arguments.
                    for (i, c) in captured.iter().enumerate() {
                        args.insert(i, Argument::positional(Expression::ident(c)));
                    }
                }
            }
        }
        ExprKind::Ident(n) if n == old && !other_candidates && captured.is_empty() => {
            *expr = Expression::ident(&sig.mangled);
        }
        _ => {}
    }
}

/// `ovl_sig_accepts` with the first `skip` parameters (threaded captures)
/// exempt from checking.
fn ovl_sig_accepts_offset(
    sig: &OvlSig,
    args: &[Argument],
    arg_tys: &[Option<OvlTy>],
    skip: usize,
) -> bool {
    arg_tys.iter().enumerate().all(|(i, ty)| {
        let Some(t) = ty else { return true };
        let pi = i + skip;
        let Some(p) = sig.param_tys.get(pi) else { return sig.has_rest };
        let nullable = sig.param_nullable.get(pi).copied().unwrap_or(false);
        if *t == OvlTy::Other {
            if let Some(arg_cls) = args.get(i).and_then(|a| ovl_arg_class_name(&a.value)) {
                let param_cls = sig.param_class_names.get(pi).and_then(|c| c.as_deref());
                return match param_cls {
                    Some(pc) => pc == arg_cls || pc == "Any" || is_user_class_name(pc),
                    None => true };
            }
        }
        match t {
            OvlTy::Null => nullable,
            OvlTy::Int => matches!(p, OvlTy::Int | OvlTy::Float | OvlTy::Other),
            _ => *p == *t || *p == OvlTy::Other }
    })
}

/// Coarse type classes for overload resolution — just enough to tell the
/// common overload pairs apart at a CALL SITE with literal/lambda arguments.
#[derive(Debug, Clone, Copy, PartialEq)]
enum OvlTy {
    Int,
    Float,
    Str,
    Bool,
    List,
    Set,
    Lambda,
    Pair,
    Null,
    Other,
}

fn ovl_param_ty(hint: Option<&str>) -> OvlTy {
    let Some(hint) = hint else { return OvlTy::Other };
    let bare = hint
        .trim()
        .trim_end_matches('?')
        .split('<')
        .next()
        .unwrap_or(hint)
        .trim();
    if hint.contains("->") {
        return OvlTy::Lambda;
    }
    match bare {
        "Int" | "Long" | "Short" | "Byte" => OvlTy::Int,
        "Double" | "Float" => OvlTy::Float,
        "String" | "Char" | "CharSequence" => OvlTy::Str,
        "Boolean" => OvlTy::Bool,
        "Set" | "MutableSet" | "HashSet" | "LinkedHashSet" => OvlTy::Set,
        "List" | "MutableList" | "Array" | "IntArray" | "Collection" | "Iterable" => {
            OvlTy::List
        }
        "Pair" => OvlTy::Pair,
        _ => OvlTy::Other }
}

/// `Pair<Int, …>` → the FIRST type argument's class (tells same-arity Pair
/// overloads apart).
fn ovl_param_pair_first(hint: Option<&str>) -> Option<OvlTy> {
    let hint = hint?;
    let inner = hint.split_once('<')?.1;
    let first = inner.split([',', '>']).next()?.trim();
    if first.is_empty() {
        return None;
    }
    Some(ovl_param_ty(Some(first)))
}

/// A Pair-valued ARGUMENT's first element class.
fn ovl_arg_pair_first(expr: &Expression) -> Option<OvlTy> {
    match &expr.kind {
        ExprKind::Tuple(items) => items.first().and_then(ovl_arg_ty),
        ExprKind::Call { callee, args, .. }
            if matches!(&callee.kind, ExprKind::Ident(n) if n == "Pair") =>
        {
            args.first().and_then(|a| ovl_arg_ty(&a.value))
        }
        _ => None }
}

/// A user-class-typed argument's class name (`Child()` → "Child").
fn ovl_arg_class_name(expr: &Expression) -> Option<&str> {
    match &expr.kind {
        ExprKind::New { class, .. } => match &class.kind {
            ExprKind::Ident(n) => Some(n.as_str()),
            _ => None },
        ExprKind::Call { callee, .. } => match &callee.kind {
            ExprKind::Ident(n) if is_user_class_name(n) => Some(n.as_str()),
            _ => None },
        _ => None }
}

fn ovl_arg_ty(expr: &Expression) -> Option<OvlTy> {
    match &expr.kind {
        ExprKind::New { .. } => Some(OvlTy::Other),
        ExprKind::Lit(Literal::Int(_)) => Some(OvlTy::Int),
        ExprKind::Lit(Literal::Float(_)) => Some(OvlTy::Float),
        ExprKind::Lit(Literal::Str(_)) => Some(OvlTy::Str),
        ExprKind::Lit(Literal::Bool(_)) => Some(OvlTy::Bool),
        ExprKind::Lit(Literal::Null) => Some(OvlTy::Null),
        ExprKind::Lambda { .. } => Some(OvlTy::Lambda),
        ExprKind::Tuple(_) => Some(OvlTy::Pair),
        ExprKind::Array(_) => Some(OvlTy::List),
        // Set literals are already lowered to marker-carrying dict objects.
        ExprKind::Object(props) => {
            let is_set = props.iter().any(|p| match p {
                ObjectProperty::KeyValue { key, .. } => {
                    matches!(&key.kind, ExprKind::Lit(Literal::Str(s)) if s == "__kt_set")
                }
                _ => false });
            if is_set { Some(OvlTy::Set) } else { Some(OvlTy::Other) }
        }
        ExprKind::Call { callee, .. } => match &callee.kind {
            ExprKind::Ident(n) if matches!(n.as_str(), "setOf" | "mutableSetOf") => {
                Some(OvlTy::Set)
            }
            ExprKind::Ident(n) if matches!(n.as_str(), "listOf" | "mutableListOf" | "arrayOf"
                | "intArrayOf") => Some(OvlTy::List),
            ExprKind::Ident(n) if n == "Pair" => Some(OvlTy::Pair),
            ExprKind::Member { field, .. } if field == "toString" => Some(OvlTy::Str),
            ExprKind::Member { field, .. } if matches!(field.as_str(), "toInt" | "count") => {
                Some(OvlTy::Int)
            }
            ExprKind::Ident(n) if is_user_class_name(n) => Some(OvlTy::Other),
            _ => None },
        ExprKind::Binary { op, left, right } => match op {
            BinOp::Add | BinOp::Sub | BinOp::Mul | BinOp::Div | BinOp::Mod => {
                match (ovl_arg_ty(left), ovl_arg_ty(right)) {
                    (Some(OvlTy::Str), _) | (_, Some(OvlTy::Str)) if *op == BinOp::Add => {
                        Some(OvlTy::Str)
                    }
                    (Some(OvlTy::Float), _) | (_, Some(OvlTy::Float)) => Some(OvlTy::Float),
                    (Some(OvlTy::Int), Some(OvlTy::Int)) => Some(OvlTy::Int),
                    _ => None }
            }
            BinOp::Eq | BinOp::NotEq | BinOp::Lt | BinOp::Gt | BinOp::LtEq | BinOp::GtEq => {
                Some(OvlTy::Bool)
            }
            _ => None },
        _ => None }
}

#[derive(Debug, Clone)]
struct OvlSig {
    mangled: String,
    min_args: usize,
    max_args: usize,
    param_tys: Vec<OvlTy>,
    param_nullable: Vec<bool>,
    param_pair_first: Vec<Option<OvlTy>>,
    param_class_names: Vec<Option<String>>,
    has_rest: bool,
}

/// Top-level `fun` OVERLOADS: the shared module has one function per name,
/// so same-name declarations silently replaced each other (last won). Each
/// overload gets a mangled name and every visible call site is resolved
/// statically — by arity first, then by literal/lambda argument classes.
fn kotlin_resolve_function_overloads(body: &mut Vec<Statement>) {
    use std::collections::HashMap as Map;
    let mut groups: Map<String, Vec<usize>> = Map::new();
    for (i, stmt) in body.iter().enumerate() {
        if let StmtKind::FunctionDecl { name, .. } = &stmt.kind {
            if name != "main" {
                groups.entry(name.clone()).or_default().push(i);
            }
        }
    }
    // Local `fun`s that reuse a top-level name: the emitted chunk table is
    // flat, so the local silently REPLACED the global. Those top-level
    // singletons join the table (getting a mangled name) so calls can pick.
    let mut local_fn_names: HashSet<String> = HashSet::new();
    fn collect_local_fn_names(stmts: &[Statement], top: bool, out: &mut HashSet<String>) {
        for s in stmts {
            if let StmtKind::FunctionDecl { name, body, .. } = &s.kind {
                if !top {
                    out.insert(name.clone());
                }
                collect_local_fn_names(body, false, out);
            }
        }
    }
    collect_local_fn_names(body, true, &mut local_fn_names);
    let mut table: Map<String, Vec<OvlSig>> = Map::new();
    for (name, indices) in groups {
        // Singletons join when a LOCAL fn or a CLASS METHOD shares the name
        // — the flat chunk table would let one silently replace the other.
        let class_shares_name = (0..=6).any(|a| is_user_member_name(&name, a));
        if indices.len() < 2 && !local_fn_names.contains(&name) && !class_shares_name {
            continue;
        }
        let mut sigs = Vec::new();
        for (k, &i) in indices.iter().enumerate() {
            if let StmtKind::FunctionDecl { name: decl_name, params, .. } = &mut body[i].kind {
                let mangled = format!("{name}__ovl{k}");
                let has_rest = params.iter().any(|p| p.is_rest);
                let max = if has_rest { usize::MAX } else { params.len() };
                let min = params
                    .iter()
                    .filter(|p| p.default.is_none() && !p.is_rest)
                    .count();
                let param_tys = params
                    .iter()
                    .map(|p| ovl_param_ty(p.type_hint.as_ref().map(|t| t.spelling())))
                    .collect();
                let param_nullable = params
                    .iter()
                    .map(|p| {
                        p.is_nullable
                            || p.type_hint
                                .as_ref()
                                .is_some_and(|t| t.spelling().trim_end().ends_with('?'))
                    })
                    .collect();
                // Generic arguments only survive in the RAW recording — the
                // TypeHint spelling is stripped to the bare head.
                let raw = KOTLIN_FN_RAW_PARAM_TYPES.with(|map| {
                    map.borrow().get(&name).and_then(|v| v.get(k).cloned())
                });
                let param_pair_first = (0..params.len())
                    .map(|pi| {
                        raw.as_ref()
                            .and_then(|r| r.get(pi))
                            .and_then(|t| ovl_param_pair_first(Some(t.as_str())))
                            .or_else(|| {
                                ovl_param_pair_first(
                                    params[pi].type_hint.as_ref().map(|t| t.spelling()),
                                )
                            })
                    })
                    .collect();
                let param_class_names = params
                    .iter()
                    .map(|p| {
                        p.type_hint
                            .as_ref()
                            .map(|t| t.spelling().trim().trim_end_matches('?').to_string())
                    })
                    .collect();
                *decl_name = mangled.clone();
                sigs.push(OvlSig {
                    mangled,
                    min_args: min,
                    max_args: max,
                    param_tys,
                    param_nullable,
                    param_pair_first,
                    param_class_names,
                    has_rest,
                });
            }
        }
        table.insert(name, sigs);
    }
    if table.is_empty() {
        return;
    }
    let mut shadowed: HashMap<String, OvlSig> = HashMap::new();
    for stmt in body.iter_mut() {
        kotlin_rewrite_overload_calls_stmt(stmt, &table, &mut shadowed);
    }
}

fn ovl_sig_of_params(name: &str, params: &[Param]) -> OvlSig {
    let has_rest = params.iter().any(|p| p.is_rest);
    OvlSig {
        mangled: name.to_string(),
        min_args: params
            .iter()
            .filter(|p| p.default.is_none() && !p.is_rest)
            .count(),
        max_args: if has_rest { usize::MAX } else { params.len() },
        param_tys: params
            .iter()
            .map(|p| ovl_param_ty(p.type_hint.as_ref().map(|t| t.spelling())))
            .collect(),
        param_nullable: params.iter().map(|p| p.is_nullable).collect(),
        param_pair_first: params
            .iter()
            .map(|p| ovl_param_pair_first(p.type_hint.as_ref().map(|t| t.spelling())))
            .collect(),
        param_class_names: params
            .iter()
            .map(|p| p.type_hint.as_ref().map(|t| t.spelling().trim().to_string()))
            .collect(),
        has_rest,
    }
}

fn ovl_sig_accepts(
    sig: &OvlSig,
    args: &[Argument],
    arg_tys: &[Option<OvlTy>],
    strict: bool,
) -> bool {
    arg_tys.iter().enumerate().all(|(i, ty)| {
        let Some(t) = ty else { return true };
        // Vararg tail: every extra argument checks against the rest param.
        let pi = if sig.has_rest && i >= sig.param_tys.len().saturating_sub(1) {
            sig.param_tys.len().saturating_sub(1)
        } else {
            i
        };
        let (Some(p), nullable) = (
            sig.param_tys.get(pi),
            sig.param_nullable.get(pi).copied().unwrap_or(false),
        ) else {
            return false;
        };
        // USER-class arguments (`visit(Child())`): exact class name in the
        // strict pass; ANY user-class param accepts in the widening pass
        // (the superclass overload catches the rest).
        if *t == OvlTy::Other {
            if let Some(arg_cls) = args.get(i).and_then(|a| ovl_arg_class_name(&a.value)) {
                let param_cls = sig.param_class_names.get(pi).and_then(|c| c.as_deref());
                return match param_cls {
                    Some(pc) if strict => pc == arg_cls,
                    Some(pc) => {
                        pc == arg_cls || pc == "Any" || is_user_class_name(pc)
                    }
                    None => !strict };
            }
        }
        // Same-arity `Pair<Int,…>` vs `Pair<String,…>` overloads compare
        // the FIRST type argument when both sides expose one.
        if *t == OvlTy::Pair && *p == OvlTy::Pair {
            if let (Some(Some(want)), Some(have)) = (
                sig.param_pair_first.get(pi),
                args.get(i).and_then(|a| ovl_arg_pair_first(&a.value)),
            ) {
                return *want == have;
            }
        }
        match t {
            OvlTy::Null => nullable,
            OvlTy::Int if !strict => {
                matches!(p, OvlTy::Int | OvlTy::Float | OvlTy::Other)
            }
            _ if !strict => *p == *t || *p == OvlTy::Other,
            _ => *p == *t }
    })
}

fn kotlin_pick_overload<'t>(sigs: &'t [OvlSig], args: &[Argument]) -> Option<&'t OvlSig> {
    let argc = args.len();
    let by_arity: Vec<&OvlSig> = sigs
        .iter()
        .filter(|s| s.min_args <= argc && argc <= s.max_args)
        .collect();
    match by_arity.len() {
        0 => sigs.first(),
        1 => Some(by_arity[0]),
        _ => {
            let arg_tys: Vec<Option<OvlTy>> =
                args.iter().map(|a| ovl_arg_ty(&a.value)).collect();
            // Kotlin's preference order: exact classes without defaults or
            // varargs first, then with defaults, then widening, varargs last.
            let passes: [(bool, bool, bool); 4] = [
                (true, false, false),  // strict, no defaults used, no rest
                (true, true, false),   // strict, defaults allowed, no rest
                (false, true, false),  // widening, no rest
                (false, true, true),   // anything incl. rest
            ];
            for (strict, allow_defaults, allow_rest) in passes {
                if let Some(s) = by_arity.iter().find(|s| {
                    (!s.has_rest || allow_rest)
                        && (allow_defaults || (!s.has_rest && s.max_args == argc))
                        && ovl_sig_accepts(s, args, &arg_tys, strict)
                }) {
                    return Some(s);
                }
            }
            by_arity.first().copied()
        }
    }
}

fn kotlin_rewrite_overload_calls_stmt(
    stmt: &mut Statement,
    table: &HashMap<String, Vec<OvlSig>>,
    shadowed: &mut HashMap<String, OvlSig>,
) {
    // A LOCAL `fun` with an overloaded top-level name shadows the group —
    // but only for calls its OWN signature accepts; Kotlin still resolves
    // `ping(1)` to the global when the local takes a String.
    if let StmtKind::FunctionDecl { name, params, body, .. } = &mut stmt.kind {
        // Rename shadowing locals first — the flat chunk table would let
        // them replace the global otherwise.
        let mut locally_declared: Vec<(String, OvlSig)> = Vec::new();
        for s in body.iter_mut() {
            if let StmtKind::FunctionDecl { name: local_name, params: local_params, .. } =
                &mut s.kind
                && table.contains_key(local_name.as_str())
            {
                let renamed = format!("{local_name}__local");
                let mut sig = ovl_sig_of_params(local_name, local_params);
                sig.mangled = renamed.clone();
                locally_declared.push((local_name.clone(), sig));
                *local_name = renamed;
            }
        }
        let self_shadow = table
            .contains_key(name.as_str())
            .then(|| (name.clone(), ovl_sig_of_params(name, params)));
        let mut restored = Vec::new();
        for (n, sig) in locally_declared.into_iter().chain(self_shadow) {
            restored.push((n.clone(), shadowed.insert(n, sig)));
        }
        for s in body.iter_mut() {
            kotlin_rewrite_overload_calls_stmt(s, table, shadowed);
        }
        for (n, prev) in restored {
            match prev {
                Some(sig) => {
                    shadowed.insert(n, sig);
                }
                None => {
                    shadowed.remove(&n);
                }
            }
        }
        return;
    }
    stmt.walk_exprs_mut(&mut |expr| {
        kotlin_rewrite_overload_calls_expr(expr, table, shadowed);
    });
}

fn kotlin_rewrite_overload_calls_expr(
    expr: &mut Expression,
    table: &HashMap<String, Vec<OvlSig>>,
    shadowed: &HashMap<String, OvlSig>,
) {
    if let ExprKind::Call { callee, args, .. } = &mut expr.kind
        && let ExprKind::Ident(name) = &callee.kind
        && let Some(sigs) = table.get(name.as_str())
    {
        if let Some(local) = shadowed.get(name.as_str()) {
            let arg_tys: Vec<Option<OvlTy>> =
                args.iter().map(|a| ovl_arg_ty(&a.value)).collect();
            let argc = args.len();
            if local.min_args <= argc
                && argc <= local.max_args
                && ovl_sig_accepts(local, args, &arg_tys, false)
            {
                let mangled = local.mangled.clone();
                **callee = Expression::ident(&mangled);
                return; // the local wins
            }
        }
        if let Some(sig) = kotlin_pick_overload(sigs, args) {
            **callee = Expression::ident(&sig.mangled);
        }
    }
}

fn kotlin_import_aliases(imports: &[Import]) -> HashMap<String, String> {
    let mut aliases = HashMap::new();
    for import in imports {
        match &import.kind {
            ImportKind::Simple {
                path,
                alias: Some(alias) } => {
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

fn rewrite_imported_value_ident(
    name: &mut String,
    aliases: &HashMap<String, String>,
    scope: &HashSet<String>,
) {
    if scope.contains(name) {
        return;
    }
    if let Some(path) = aliases.get(name) {
        if let Some(leaf) = imported_leaf(path) {
            *name = leaf.to_string();
        }
    }
}

fn kotlin_import_path(path: &str) -> String {
    let Some(leaf) = path.rsplit('.').next() else {
        return path.to_string();
    };
    let java_util = match leaf {
        "ArrayList" | "HashMap" | "HashSet" | "LinkedHashMap" | "LinkedHashSet" => Some(leaf),
        _ => None };
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
            else_body } => {
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
            else_body } => {
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
            body } => {
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
            finally } => {
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
            if let ExprKind::Ident(name) = &mut object.kind {
                rewrite_imported_value_ident(name, aliases, scope);
            } else {
                rewrite_import_aliases_in_expr(object, aliases, scope);
            }
            rewrite_imported_value_ident(field, aliases, scope);
        }
        ExprKind::Call { callee, args, .. } => {
            if let ExprKind::Ident(name) = &mut callee.kind {
                rewrite_imported_value_ident(name, aliases, scope);
            } else {
                rewrite_import_aliases_in_expr(callee, aliases, scope);
            }
            for arg in args {
                rewrite_import_aliases_in_expr(&mut arg.value, aliases, scope);
            }
        }
        ExprKind::New { class, args } => {
            if let ExprKind::Ident(name) = &mut class.kind {
                rewrite_imported_value_ident(name, aliases, scope);
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
                optional: false }))
        }
        ExprKind::Call { callee, args, .. } => match &callee.kind {
            ExprKind::Ident(name)
                if matches!(name.as_str(), "compareBy" | "compareByDescending")
                    && args.len() == 1 =>
            {
                Some(kotlin_compare_by_lambda(
                    args[0].value.clone(),
                    name == "compareByDescending",
                ))
            }
            // `cmp.thenBy { sel }` — tie-break combinator over ANY comparator
            // expression (comparators here are plain 2-arg lambdas).
            ExprKind::Member { object, field, .. }
                if matches!(field.as_str(), "thenBy" | "thenByDescending")
                    && args.len() == 1 =>
            {
                Some(kotlin_then_by_lambda(
                    (**object).clone(),
                    args[0].value.clone(),
                    field == "thenByDescending",
                ))
            }
            ExprKind::Ident(name)
                if matches!(
                    name.as_str(),
                    "listOf"
                        | "mutableListOf"
                        | "arrayListOf"
                        | "arrayOf"
                        | "emptyList"
                        | "intArrayOf"
                        | "doubleArrayOf"
                        | "booleanArrayOf"
                        | "charArrayOf"
                        | "longArrayOf"
                        | "sequenceOf"
                ) =>
            {
                Some(Expression::new(ExprKind::Array(
                    args.iter()
                        .map(|arg| ArrayElement {
                            key: None,
                            value: arg.value.clone(),
                            spread: false,
                            by_ref: false })
                        .collect(),
                )))
            }
            ExprKind::Ident(name)
                if matches!(name.as_str(), "buildList" | "buildSet" | "buildMap")
                    && !args.is_empty()
                    && matches!(args.last().unwrap().value.kind, ExprKind::Lambda { .. }) =>
            {
                // `buildList(capacity) { }` — the capacity is a growth HINT,
                // not a size; ignored (the 2-arg form built nothing).
                kotlin_build_collection_expr(name, &args.last().unwrap().value)
            }
            ExprKind::Ident(name)
                if matches!(
                    name.as_str(),
                    "setOf"
                        | "mutableSetOf"
                        | "linkedSetOf"
                        | "hashSetOf"
                        | "buildSet"
                        | "emptySet"
                ) =>
            {
                Some(create_kotlin_set_expr(
                    args.iter().map(|arg| arg.value.clone()).collect(),
                ))
            }
            ExprKind::Ident(name)
                if matches!(name.as_str(), "buildList" | "buildSet" | "buildMap")
                    && args.len() == 1 =>
            {
                kotlin_build_collection_expr(name, &args[0].value)
            }
            // Only the LAMBDA form is the scope function — a user `fun run(x)`
            // called with a value must stay an ordinary call ("f64 is not
            // callable" when the argument was invoked).
            ExprKind::Ident(name)
                if name == "run"
                    && args.len() == 1
                    && matches!(args[0].value.kind, ExprKind::Lambda { .. }) =>
            {
                Some(Expression::new(ExprKind::Call {
                    callee: Box::new(args[0].value.clone()),
                    args: Vec::new(),
                    optional: false }))
            }
            ExprKind::Ident(name) if name == "joinToString" && !args.is_empty() => {
                let (items, separator) = if args.len() >= 2 {
                    let mapped = kotlin_map_call_expr(args[0].value.clone(), args[1].clone());
                    (mapped, args[0].value.clone())
                } else {
                    (
                        args[0].value.clone(),
                        Expression::new(ExprKind::Lit(Literal::Str(", ".into()))),
                    )
                };
                Some(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__coll_join")),
                    args: vec![Argument::positional(items), Argument::positional(separator)],
                    optional: false }))
            }
            ExprKind::Member { object, field, .. } if field == "joinToString" => {
                let object = kotlin_materialize_generate_sequence(object, None)
                    .unwrap_or_else(|| *object.clone());
                let (items, separator) = if args.len() >= 2 {
                    if let Some(static_items) = kotlin_static_array_items(&object) {
                        if let Some(mapped) =
                            kotlin_apply_static_join_transform(&static_items, &args[1].value)
                        {
                            (mapped, args[0].value.clone())
                        } else {
                            let mapped = kotlin_map_call_expr(object.clone(), args[1].clone());
                            (mapped, args[0].value.clone())
                        }
                    } else {
                        let mapped = kotlin_map_call_expr(object.clone(), args[1].clone());
                        (mapped, args[0].value.clone())
                    }
                } else {
                    (
                        object,
                        args.first()
                            .map(|arg| arg.value.clone())
                            .unwrap_or_else(|| {
                                Expression::new(ExprKind::Lit(Literal::Str(", ".into())))
                            }),
                    )
                };
                Some(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__coll_join")),
                    args: vec![Argument::positional(items), Argument::positional(separator)],
                    optional: false }))
            }
            // `sortedBy` stays a member call: `[array_methods]` routes it to
            // `__array_sort_by_key` (1-arg KEY SELECTOR). It used to be mapped
            // onto `sort`, whose lambda is a 2-arg COMPARATOR — the key values
            // were compared as if they were compare results, reversing runs.
            ExprKind::Member { object, field, null_safe } if field == "let" && args.len() == 1 => {
                let call = Expression::new(ExprKind::Call {
                    callee: Box::new(args[0].value.clone()),
                    args: vec![Argument::positional(*object.clone())],
                    optional: false });
                // `x?.let { }` guards the receiver; the guarded shape is
                // also what the `?: fallback` rewrite below recognizes.
                Some(if *null_safe {
                    Expression::new(ExprKind::Ternary {
                        cond: Box::new(Expression::new(ExprKind::Binary {
                            op: BinOp::Eq,
                            left: object.clone(),
                            right: Box::new(Expression::null()) })),
                        then: Box::new(Expression::null()),
                        else_: Box::new(call) })
                } else {
                    call
                })
            }
            // The scope-function family. `let` above passes the receiver as
            // `it`; these four wrap it in an IIFE with the right RETURN and
            // — for the receiver-lambdas `apply`/`run` — bind bare member
            // references onto the receiver.
            ExprKind::Member { object, field, .. }
                if matches!(field.as_str(), "also" | "takeIf" | "takeUnless")
                    && args.len() == 1
                    && matches!(args[0].value.kind, ExprKind::Lambda { .. })
                    && !is_user_member_name(field, 1) =>
            {
                Some(kotlin_it_scope_iife(
                    (**object).clone(),
                    args[0].value.clone(),
                    field,
                ))
            }
            ExprKind::Member { object, field, .. }
                if matches!(field.as_str(), "apply" | "run")
                    && args.len() == 1
                    && matches!(args[0].value.kind, ExprKind::Lambda { .. })
                    && !is_user_member_name(field, 1) =>
            {
                kotlin_receiver_scope_iife(
                    (**object).clone(),
                    &args[0].value,
                    /*returns_receiver:*/ field == "apply",
                )
            }
            // `with(x) { body }` is `x.run { body }`.
            ExprKind::Ident(name)
                if name == "with"
                    && args.len() == 2
                    && matches!(args[1].value.kind, ExprKind::Lambda { .. }) =>
            {
                kotlin_receiver_scope_iife(
                    args[0].value.clone(),
                    &args[1].value,
                    /*returns_receiver:*/ false,
                )
            }
            _ => None },
        _ => None }
}

/// `x.also { f }` / `x.takeIf { p }` / `x.takeUnless { p }` — the receiver
/// flows through as the user lambda's `it`; only the RETURN differs.
///
/// A CHEAP receiver (ident/literal) skips the wrapper lambda: writes to
/// captured outer vars only propagate ONE lambda level up, so nesting the
/// user lambda inside a wrapper lost `takeIf { checks++; … }` increments.
fn kotlin_it_scope_iife(recv: Expression, lambda: Expression, kind: &str) -> Expression {
    if matches!(recv.kind, ExprKind::Ident(_) | ExprKind::Lit(_)) {
        let call_lambda = Expression::new(ExprKind::Call {
            callee: Box::new(lambda),
            args: vec![Argument::positional(recv.clone())],
            optional: false });
        return match kind {
            // `(f(x), x)` — comma sequence keeps the receiver as the value.
            "also" => Expression::new(ExprKind::Sequence(vec![call_lambda, recv])),
            _ => {
                let (then, else_) = if kind == "takeIf" {
                    (recv.clone(), Expression::null())
                } else {
                    (Expression::null(), recv.clone())
                };
                Expression::new(ExprKind::Ternary {
                    cond: Box::new(call_lambda),
                    then: Box::new(then),
                    else_: Box::new(else_) })
            }
        };
    }
    // Non-cheap receiver: INLINE the user lambda's body into the wrapper
    // (calling it as a nested function would put captured-var WRITES two
    // lambda levels deep, where they do not propagate). The wrapper reuses
    // the user lambda's own parameter name, so the body needs no renaming.
    let (param_name, mut stmts) = match &lambda.kind {
        ExprKind::Lambda { params, body, .. } => {
            let name = params
                .first()
                .map(|p| p.name.clone())
                .unwrap_or_else(|| "it".to_string());
            let stmts = match body {
                LambdaBody::Block(stmts) => stmts.clone(),
                LambdaBody::Expr(e) => {
                    vec![Statement::new(StmtKind::Expr((**e).clone()))]
                }
            };
            (name, stmts)
        }
        _ => {
            // Not a literal lambda — fall back to calling the value.
            let r = "__kt_scope_r".to_string();
            let call = Expression::new(ExprKind::Call {
                callee: Box::new(lambda.clone()),
                args: vec![Argument::positional(Expression::ident(&r))],
                optional: false });
            (r, vec![Statement::new(StmtKind::Expr(call))])
        }
    };
    match kind {
        "also" => {
            // The body's value is discarded; the receiver comes back.
            if let Some(last) = stmts.pop() {
                match last.kind {
                    StmtKind::Return(Some(e)) => {
                        stmts.push(Statement::new(StmtKind::Expr(e)))
                    }
                    other => stmts.push(Statement::new(other)) }
            }
            stmts.push(Statement::new(StmtKind::Return(Some(Expression::ident(
                &param_name,
            )))));
        }
        _ => {
            // The trailing value is the PREDICATE.
            kotlin_return_wrap_last(&mut stmts);
            let (then, else_) = if kind == "takeIf" {
                (Expression::ident(&param_name), Expression::null())
            } else {
                (Expression::null(), Expression::ident(&param_name))
            };
            if let Some(last) = stmts.pop() {
                match last.kind {
                    StmtKind::Return(Some(pred)) => {
                        stmts.push(Statement::new(StmtKind::Return(Some(Expression::new(
                            ExprKind::Ternary {
                                cond: Box::new(pred),
                                then: Box::new(then),
                                else_: Box::new(else_) },
                        )))));
                    }
                    other => stmts.push(Statement::new(other)) }
            }
        }
    }
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: vec![kt_param(&param_name)],
            body: LambdaBody::Block(stmts),
            captures: Vec::new(),
            is_async: false })),
        args: vec![Argument::positional(recv)],
        optional: false })
}

/// `x.apply { body }` / `x.run { body }` — RECEIVER lambdas: bare property
/// reads/writes and bare method calls bind to the receiver, `this` IS the
/// receiver. Lowered to an IIFE whose parameter takes the receiver's place.
fn kotlin_receiver_scope_iife(
    recv: Expression,
    lambda: &Expression,
    returns_receiver: bool,
) -> Option<Expression> {
    let ExprKind::Lambda { params, body, .. } = &lambda.kind else {
        return None;
    };
    // Receiver lambdas are parameter-less; the implicit `it` is the walker's.
    if params.len() > 1 || params.first().is_some_and(|p| p.name != "it") {
        return None;
    }
    let r = "__kt_builder";
    let mut stmts: Vec<Statement> = match body {
        LambdaBody::Block(stmts) => stmts.clone(),
        LambdaBody::Expr(expr) => vec![Statement::new(StmtKind::Expr((**expr).clone()))] };
    // Names the lambda body DECLARES itself shadow the receiver's members
    // (`val label = "with"` inside `with(node) { }`).
    let mut body_locals = HashSet::new();
    kotlin_local_binding_names(&stmts, &mut body_locals);
    for stmt in &mut stmts {
        // Receiver refs FIRST: a user-declared `add` must become a member
        // call before the builder table reads it as the collection mutator.
        kotlin_rewrite_receiver_refs_stmt(stmt, r, &body_locals);
        // Bare collection mutators (`add`, `put`, …) onto the receiver —
        // the same lowering builder lambdas use (the param name matches).
        kotlin_rewrite_builder_calls_stmt(stmt);
    }
    if returns_receiver {
        // The body's value is discarded; the receiver comes back.
        if let Some(last) = stmts.pop() {
            match last.kind {
                StmtKind::Return(Some(e)) => stmts.push(Statement::new(StmtKind::Expr(e))),
                other => stmts.push(Statement::new(other)) }
        }
        stmts.push(Statement::new(StmtKind::Return(Some(Expression::ident(r)))));
    } else {
        kotlin_return_wrap_last(&mut stmts);
    }
    // Keep the receiver's STATIC type on the parameter: StringBuilder's
    // (and other tree-dispatched) members resolve by declared type, and an
    // untyped IIFE param answered undefined for `append`.
    let mut param = kt_param(r);
    let empty_locals = KotlinLocalTypes::new();
    let empty_ops = KotlinOperatorTable::new();
    if let Some(ty) = kotlin_expr_type(&recv, &empty_locals, &empty_ops) {
        param.type_hint = Some(ty.into());
    }
    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: vec![param],
            body: LambdaBody::Block(stmts),
            captures: Vec::new(),
            is_async: false })),
        args: vec![Argument::positional(recv)],
        optional: false }))
}

/// Bind a receiver-lambda body's bare member references onto `recv`:
/// `this` → the receiver; a bare `Ident` that names a source-declared
/// PROPERTY → `recv.prop`; a bare call of a source-declared METHOD →
/// `recv.method(…)`. Everything else (locals, top-level functions) passes
/// through untouched.
fn kotlin_rewrite_receiver_refs_stmt(stmt: &mut Statement, recv: &str, skip: &HashSet<String>) {
    match &mut stmt.kind {
        StmtKind::Expr(e) | StmtKind::Return(Some(e)) | StmtKind::Throw { expr: Some(e), .. } => {
            kotlin_rewrite_receiver_refs_expr(e, recv, skip)
        }
        StmtKind::VarDecl { declarations, .. } => {
            for d in declarations {
                if let Some(init) = &mut d.init {
                    kotlin_rewrite_receiver_refs_expr(init, recv, skip);
                }
            }
        }
        StmtKind::If { cond, then_body, elifs, else_body } => {
            kotlin_rewrite_receiver_refs_expr(cond, recv, skip);
            for s in then_body {
                kotlin_rewrite_receiver_refs_stmt(s, recv, skip);
            }
            for (c, b) in elifs {
                kotlin_rewrite_receiver_refs_expr(c, recv, skip);
                for s in b {
                    kotlin_rewrite_receiver_refs_stmt(s, recv, skip);
                }
            }
            if let Some(b) = else_body {
                for s in b {
                    kotlin_rewrite_receiver_refs_stmt(s, recv, skip);
                }
            }
        }
        StmtKind::While { cond, body, .. } => {
            kotlin_rewrite_receiver_refs_expr(cond, recv, skip);
            for s in body {
                kotlin_rewrite_receiver_refs_stmt(s, recv, skip);
            }
        }
        StmtKind::ForIn { iter, body, .. } => {
            kotlin_rewrite_receiver_refs_expr(iter, recv, skip);
            for s in body {
                kotlin_rewrite_receiver_refs_stmt(s, recv, skip);
            }
        }
        _ => {}
    }
}

fn kotlin_rewrite_receiver_refs_expr(expr: &mut Expression, recv: &str, skip: &HashSet<String>) {
    match &mut expr.kind {
        ExprKind::This => {
            *expr = Expression::ident(recv);
            return;
        }
        ExprKind::Ident(name) => {
            if skip.contains(name.as_str()) {
                return;
            }
            // A bare Member built HERE runs after the walk-time member
            // lowering, so the probing spellings must be produced directly:
            // `second` inside `fun Pair<..>.delta()` is a tuple slot, `keys`
            // inside a Map extension is the dict view.
            if matches!(name.as_str(), "first" | "second" | "third")
                && !USER_PROPERTY_NAMES.with(|set| set.borrow().contains(name.as_str()))
            {
                let idx = match name.as_str() {
                    "first" => 0,
                    "second" => 1,
                    _ => 2 };
                *expr = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__kt_tuple_prop")),
                    args: vec![
                        Argument::positional(Expression::ident(recv)),
                        Argument::positional(Expression::new(ExprKind::Lit(Literal::Str(
                            name.clone().into(),
                        )))),
                        Argument::positional(Expression::int(idx)),
                    ],
                    optional: false });
                return;
            }
            if matches!(name.as_str(), "keys" | "values" | "entries")
                && !USER_PROPERTY_NAMES.with(|set| set.borrow().contains(name.as_str()))
            {
                let helper = match name.as_str() {
                    "keys" => "__dict_keys",
                    "values" => "__dict_values",
                    _ => "__kt_map_entries" };
                *expr = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident(helper)),
                    args: vec![Argument::positional(Expression::ident(recv))],
                    optional: false });
                return;
            }
            // Source-declared properties, plus the universal receiver
            // properties (`length + 1` inside `with("ok") { }`).
            if USER_PROPERTY_NAMES.with(|set| set.borrow().contains(name.as_str()))
                || matches!(
                    name.as_str(),
                    "length" | "size" | "indices" | "lastIndex"
                )
            {
                *expr = Expression::new(ExprKind::Member {
                    object: Box::new(Expression::ident(recv)),
                    field: name.clone(),
                    null_safe: false });
            }
            return;
        }
        ExprKind::Call { callee, args, .. } => {
            for a in args.iter_mut() {
                kotlin_rewrite_receiver_refs_expr(&mut a.value, recv, skip);
            }
            if let ExprKind::Ident(name) = &callee.kind {
                if skip.contains(name.as_str()) {
                    return;
                }
                // A source-declared method, or one of the common receiver
                // method spellings (StringBuilder/collection/string surface)
                // — Kotlin resolves bare calls against the lambda receiver
                // first.
                if name == "iterator" && args.is_empty() {
                    // The walk-time member pass lowers `.iterator()` — this
                    // Member is built AFTER it, so produce the lowered form
                    // directly.
                    *expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__jvm_list_iterator")),
                        args: vec![Argument::positional(Expression::ident(recv))],
                        optional: false });
                    return;
                }
                let known_receiver_method = matches!(
                    name.as_str(),
                    "append" | "appendLine" | "insert" | "delete" | "deleteCharAt"
                        | "setCharAt" | "setLength" | "reverse" | "filter" | "map"
                        | "sum" | "sorted" | "sortedBy" | "joinToString" | "take"
                        | "drop" | "uppercase" | "lowercase" | "trim" | "replace"
                        | "substring" | "split" | "count" | "average" | "forEach"
                        | "removeAt" | "clear" | "getOrDefault" | "containsKey"
                        | "toUpperCase" | "toLowerCase" | "reversed" | "capitalize"
                        | "trimIndent" | "isEmpty" | "isNotEmpty" | "padStart"
                        | "padEnd" | "indexOf" | "contains" | "first" | "last"
                );
                if is_user_member_name(name, args.len()) || known_receiver_method {
                    // Overloaded class methods dispatch by their $sig storage
                    // name — the walk-time member pass already ran, so this
                    // freshly-built member call must resolve it itself.
                    let field = overloaded_storage_name_for_args(name, args.len(), args)
                        .unwrap_or_else(|| name.clone());
                    **callee = Expression::new(ExprKind::Member {
                        object: Box::new(Expression::ident(recv)),
                        field,
                        null_safe: false });
                }
                return;
            }
            kotlin_rewrite_receiver_refs_expr(callee, recv, skip);
            return;
        }
        _ => {}
    }
    match &mut expr.kind {
        ExprKind::Binary { left, right, .. } => {
            kotlin_rewrite_receiver_refs_expr(left, recv, skip);
            kotlin_rewrite_receiver_refs_expr(right, recv, skip);
        }
        ExprKind::Assign { target, value } => {
            kotlin_rewrite_receiver_refs_expr(target, recv, skip);
            kotlin_rewrite_receiver_refs_expr(value, recv, skip);
        }
        ExprKind::Member { object, .. } => {
            kotlin_rewrite_receiver_refs_expr(object, recv, skip);
        }
        ExprKind::Index { object, index, .. } => {
            kotlin_rewrite_receiver_refs_expr(object, recv, skip);
            kotlin_rewrite_receiver_refs_expr(index, recv, skip);
        }
        ExprKind::Unary { expr: inner, .. } => {
            kotlin_rewrite_receiver_refs_expr(inner, recv, skip);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            kotlin_rewrite_receiver_refs_expr(cond, recv, skip);
            kotlin_rewrite_receiver_refs_expr(then, recv, skip);
            kotlin_rewrite_receiver_refs_expr(else_, recv, skip);
        }
        _ => {}
    }
}

/// The dotted spelling of a pure ident/member chain, if it is one.
/// `'a'..'z'` — a RANGE of one-char strings folds to the char array, so
/// `it in 'a'..'z'` and iteration both work over the materialized form.
fn kotlin_fold_char_range(expr: &mut Expression) {
    if let ExprKind::Range { start, end, inclusive } = &expr.kind {
        if let (ExprKind::Lit(Literal::Str(a)), ExprKind::Lit(Literal::Str(b))) =
            (&start.kind, &end.kind)
        {
            if a.chars().count() == 1 && b.chars().count() == 1 {
                let (a, b) = (a.chars().next().unwrap(), b.chars().next().unwrap());
                let hi = if *inclusive { b as u32 } else { b as u32 - 1 };
                let items: Vec<ArrayElement> = (a as u32..=hi)
                    .filter_map(char::from_u32)
                    .map(|c| ArrayElement {
                        key: None,
                        value: Expression::string(&c.to_string()),
                        spread: false,
                        by_ref: false })
                    .collect();
                *expr = Expression::new(ExprKind::Array(items));
            }
        }
    }
}

fn dotted_expr_path_of(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => Some(name.clone()),
        ExprKind::Member { object, field, .. } => {
            Some(format!("{}.{}", dotted_expr_path_of(object)?, field))
        }
        _ => None }
}

/// `buildList { add(1) }` / `buildSet { }` / `buildMap { put(k, v) }` — the
/// builder-receiver lambdas. Lowered to an IIFE over a fresh mutable
/// collection, with the builder's bare calls (`add`, `put`, …) rewritten onto
/// it — the same shape `run { }` lowers to, plus the receiver.
fn kotlin_build_collection_expr(kind: &str, lambda: &Expression) -> Option<Expression> {
    let factory = match kind {
        "buildList" => "mutableListOf",
        "buildSet" => "mutableSetOf",
        _ => "mutableMapOf" };
    let ExprKind::Lambda { params, body, .. } = &lambda.kind else {
        return None;
    };
    // The walker gives every lambda an implicit `it`; a builder lambda's
    // receiver is positional-less, so anything beyond that opts out.
    if params.len() > 1 || params.first().is_some_and(|p| p.name != "it") {
        return None;
    }
    let mut stmts: Vec<Statement> = match body {
        LambdaBody::Block(stmts) => stmts.clone(),
        LambdaBody::Expr(expr) => vec![Statement::new(StmtKind::Expr((**expr).clone()))] };
    for stmt in &mut stmts {
        kotlin_rewrite_builder_calls_stmt(stmt);
    }
    // The walker has already Return-wrapped the source lambda's last
    // expression; unwrap it, or the builder result below is unreachable.
    if let Some(last) = stmts.pop() {
        match last.kind {
            StmtKind::Return(Some(e)) => stmts.push(Statement::new(StmtKind::Expr(e))),
            other => stmts.push(Statement::new(other)) }
    }

    let builder = "__kt_builder";
    let mut block = vec![Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(builder.to_string()),
            type_hint: None,
            init: Some(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(factory)),
                args: Vec::new(),
                optional: false })),
            array_bounds: None,
            with_events: false }],
        kind: VarDeclKind::Let })];
    block.extend(stmts);
    block.push(Statement::new(StmtKind::Return(Some(Expression::ident(builder)))));

    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: Vec::new(),
            body: LambdaBody::Block(block),
            is_async: false,
            captures: Vec::new() })),
        args: Vec::new(),
        optional: false }))
}

/// Rewrite the builder lambda's bare mutator calls onto `__kt_builder`,
/// through the SAME `__kt_*` builtins the member spellings lower to.
fn kotlin_rewrite_builder_calls_stmt(stmt: &mut Statement) {
    match &mut stmt.kind {
        StmtKind::Expr(e) | StmtKind::Return(Some(e)) | StmtKind::Throw { expr: Some(e), .. } => {
            kotlin_rewrite_builder_calls_expr(e)
        }
        StmtKind::If { cond, then_body, elifs, else_body } => {
            kotlin_rewrite_builder_calls_expr(cond);
            for s in then_body {
                kotlin_rewrite_builder_calls_stmt(s);
            }
            for (c, b) in elifs {
                kotlin_rewrite_builder_calls_expr(c);
                for s in b {
                    kotlin_rewrite_builder_calls_stmt(s);
                }
            }
            if let Some(b) = else_body {
                for s in b {
                    kotlin_rewrite_builder_calls_stmt(s);
                }
            }
        }
        StmtKind::While { cond, body, .. } => {
            kotlin_rewrite_builder_calls_expr(cond);
            for s in body {
                kotlin_rewrite_builder_calls_stmt(s);
            }
        }
        StmtKind::ForIn { iter, body, .. } => {
            kotlin_rewrite_builder_calls_expr(iter);
            for s in body {
                kotlin_rewrite_builder_calls_stmt(s);
            }
        }
        StmtKind::For { init, cond, update, body } => {
            if let Some(s) = init {
                kotlin_rewrite_builder_calls_stmt(s);
            }
            if let Some(c) = cond {
                kotlin_rewrite_builder_calls_expr(c);
            }
            if let Some(u) = update {
                kotlin_rewrite_builder_calls_expr(u);
            }
            for s in body {
                kotlin_rewrite_builder_calls_stmt(s);
            }
        }
        StmtKind::Block(stmts) => {
            for s in stmts {
                kotlin_rewrite_builder_calls_stmt(s);
            }
        }
        StmtKind::VarDecl { declarations, .. } => {
            for d in declarations {
                if let Some(init) = &mut d.init {
                    kotlin_rewrite_builder_calls_expr(init);
                }
            }
        }
        _ => {}
    }
}

fn kotlin_rewrite_builder_calls_expr(expr: &mut Expression) {
    if let ExprKind::Call { callee, args, .. } = &mut expr.kind {
        for arg in args.iter_mut() {
            kotlin_rewrite_builder_calls_expr(&mut arg.value);
        }
        if let ExprKind::Ident(name) = &callee.kind {
            let builder_arg = Argument::positional(Expression::ident("__kt_builder"));
            let target = match (name.as_str(), args.len()) {
                ("add", 1) => Some("__kt_add"),
                ("addAll", 1) => Some("__kt_add_all"),
                ("put", 2) => Some("__dict_set"),
                ("remove", 1) => Some("__dict_delete"),
                _ => None };
            if let Some(target) = target {
                let mut new_args = vec![builder_arg];
                new_args.append(args);
                *expr = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident(target)),
                    args: new_args,
                    optional: false });
            }
        }
    }
}

fn kotlin_compare_by_lambda(selector: Expression, descending: bool) -> Expression {
    let left_name = "__kt_cmp_a";
    let right_name = "__kt_cmp_b";
    let left_key = Expression::new(ExprKind::Call {
        callee: Box::new(selector.clone()),
        args: vec![Argument::positional(Expression::ident(left_name))],
        optional: false });
    let right_key = Expression::new(ExprKind::Call {
        callee: Box::new(selector),
        args: vec![Argument::positional(Expression::ident(right_name))],
        optional: false });
    let (first, second) = if descending {
        (right_key, left_key)
    } else {
        (left_key, right_key)
    };
    let compare = Expression::new(ExprKind::Call {
        callee: Box::new(dotted_ident_expr("java.lang.Integer.compare")),
        args: vec![Argument::positional(first), Argument::positional(second)],
        optional: false });
    Expression::new(ExprKind::Lambda {
        params: vec![kt_param(left_name), kt_param(right_name)],
        body: LambdaBody::Block(vec![Statement::new(StmtKind::Return(Some(compare)))]),
        captures: Vec::new(),
        is_async: false })
}

/// `base.thenBy { sel }` → `{ a, b -> val c = base(a, b);
/// if (c != 0) c else Integer.compare(sel(a), sel(b)) }`.
fn kotlin_then_by_lambda(base: Expression, selector: Expression, descending: bool) -> Expression {
    let left_name = "__kt_cmp2_a";
    let right_name = "__kt_cmp2_b";
    let c_name = "__kt_cmp2_c";
    let base_call = Expression::new(ExprKind::Call {
        callee: Box::new(base),
        args: vec![
            Argument::positional(Expression::ident(left_name)),
            Argument::positional(Expression::ident(right_name)),
        ],
        optional: false });
    let left_key = Expression::new(ExprKind::Call {
        callee: Box::new(selector.clone()),
        args: vec![Argument::positional(Expression::ident(left_name))],
        optional: false });
    let right_key = Expression::new(ExprKind::Call {
        callee: Box::new(selector),
        args: vec![Argument::positional(Expression::ident(right_name))],
        optional: false });
    let (first, second) = if descending {
        (right_key, left_key)
    } else {
        (left_key, right_key)
    };
    let tie_break = Expression::new(ExprKind::Call {
        callee: Box::new(dotted_ident_expr("java.lang.Integer.compare")),
        args: vec![Argument::positional(first), Argument::positional(second)],
        optional: false });
    let body = vec![
        Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(c_name.to_string()),
                type_hint: None,
                init: Some(base_call),
                array_bounds: None,
                with_events: false }],
            kind: VarDeclKind::Const }),
        Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Ternary {
            cond: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::NotEq,
                left: Box::new(Expression::ident(c_name)),
                right: Box::new(Expression::int(0)) })),
            then: Box::new(Expression::ident(c_name)),
            else_: Box::new(tie_break) })))),
    ];
    Expression::new(ExprKind::Lambda {
        params: vec![kt_param(left_name), kt_param(right_name)],
        body: LambdaBody::Block(body),
        captures: Vec::new(),
        is_async: false })
}

fn kotlin_generate_sequence_take(
    sequence: &Expression,
    take_count: &Expression,
) -> Option<Expression> {
    let count = match &take_count.kind {
        ExprKind::Lit(Literal::Int(value)) if *value >= 0 => *value as usize,
        _ => return None };
    kotlin_materialize_generate_sequence(sequence, Some(count))
}

fn kotlin_sequence_source(expr: &Expression) -> Option<Expression> {
    let ExprKind::Ident(name) = &expr.kind else {
        return None;
    };
    KOTLIN_SEQUENCE_SOURCES.with(|map| map.borrow().get(name).cloned())
}

fn kotlin_is_sequence_source(expr: &Expression) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Call { callee, .. }
            if matches!(&callee.kind, ExprKind::Ident(name) if name == "generateSequence" || name == "sequence")
    )
}

fn kotlin_generate_sequence_take_while(
    sequence: &Expression,
    predicate: &Expression,
) -> Option<Expression> {
    let ExprKind::Lambda { params, body, .. } = &predicate.kind else {
        return None;
    };
    let pred_param = params.first()?.name.as_str();
    let values = kotlin_generate_sequence_values(sequence, Some(256), true)?;
    let mut out = Vec::new();
    for value in values {
        if !kotlin_eval_sequence_bool_from_value(body, pred_param, &value)? {
            break;
        }
        out.push(kotlin_seq_value_expr(&value)?);
    }
    Some(kotlin_array_expr(out))
}

#[derive(Clone)]
enum KotlinSeqValue {
    Int(i64),
    Str(String),
    Null }

fn kotlin_materialize_generate_sequence(
    sequence: &Expression,
    limit: Option<usize>,
) -> Option<Expression> {
    let values = kotlin_generate_sequence_values(sequence, limit, false)?;
    let values = values
        .into_iter()
        .filter_map(|value| kotlin_seq_value_expr(&value))
        .collect();
    Some(kotlin_array_expr(values))
}

fn kotlin_generate_sequence_values(
    sequence: &Expression,
    limit: Option<usize>,
    allow_infinite_prefix: bool,
) -> Option<Vec<KotlinSeqValue>> {
    let ExprKind::Call { callee, args, .. } = &sequence.kind else {
        return None;
    };
    if !matches!(&callee.kind, ExprKind::Ident(name) if name == "generateSequence")
        || args.len() != 2
    {
        return None;
    }
    let mut current = kotlin_seq_literal_value(&args[0].value)?;
    let ExprKind::Lambda { params, body, .. } = &args[1].value.kind else {
        return None;
    };
    let param = params.first()?.name.as_str();
    let max = limit.unwrap_or(256);
    let mut out = Vec::new();
    for _ in 0..max {
        if matches!(current, KotlinSeqValue::Null) {
            break;
        }
        out.push(current.clone());
        let next = kotlin_eval_sequence_lambda(body, param, &current)?;
        current = next;
        if limit.is_none() && matches!(current, KotlinSeqValue::Null) {
            break;
        }
    }
    if limit.is_none() && !allow_infinite_prefix && !matches!(current, KotlinSeqValue::Null) {
        return None;
    }
    Some(out)
}

fn kotlin_array_expr(values: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Array(
        values
            .into_iter()
            .map(|value| ArrayElement {
                key: None,
                value,
                spread: false,
                by_ref: false })
            .collect(),
    ))
}

fn kotlin_static_array_items(expr: &Expression) -> Option<Vec<Expression>> {
    if let Some(materialized) = kotlin_materialize_generate_sequence(expr, None) {
        return kotlin_static_array_items(&materialized);
    }
    match &expr.kind {
        ExprKind::Ident(name) => KOTLIN_STATIC_VALUES.with(|map| {
            let value = map.borrow().get(name).cloned()?;
            kotlin_static_array_items(&value)
        }),
        ExprKind::Array(items) => Some(items.iter().map(|item| item.value.clone()).collect()),
        ExprKind::Range {
            start,
            end,
            inclusive } => {
            let ExprKind::Lit(Literal::Int(start)) = start.kind else {
                return None;
            };
            let ExprKind::Lit(Literal::Int(end)) = end.kind else {
                return None;
            };
            let stop = if *inclusive { end } else { end - 1 };
            if start > stop {
                return Some(Vec::new());
            }
            Some((start..=stop).map(Expression::int).collect())
        }
        _ => None }
}

fn kotlin_chunk_static(items: &[Expression], size: usize) -> Option<Expression> {
    if size == 0 {
        return None;
    }
    let chunks = items
        .chunks(size)
        .map(|chunk| kotlin_array_expr(chunk.to_vec()))
        .collect();
    Some(kotlin_array_expr(chunks))
}

fn kotlin_window_static(items: &[Expression], size: usize) -> Option<Expression> {
    if size == 0 || items.len() < size {
        return Some(kotlin_array_expr(Vec::new()));
    }
    let windows = items
        .windows(size)
        .map(|window| kotlin_array_expr(window.to_vec()))
        .collect();
    Some(kotlin_array_expr(windows))
}

fn kotlin_zip_with_next_static(items: &[Expression]) -> Expression {
    let pairs = items
        .windows(2)
        .map(|pair| create_pair_expr(pair[0].clone(), pair[1].clone()))
        .collect();
    kotlin_array_expr(pairs)
}

fn kotlin_lambda_is_collection_sum(lambda: &Expression) -> bool {
    let ExprKind::Lambda { params, body, .. } = &lambda.kind else {
        return false;
    };
    let Some(param) = params.first() else {
        return false;
    };
    let is_sum = |expr: &Expression| {
        matches!(
            &expr.kind,
            ExprKind::Call { callee, args, .. }
                if matches!(&callee.kind, ExprKind::Ident(name) if name == "__coll_sum")
                    && args.len() == 1
                    && matches!(&args[0].value.kind, ExprKind::Ident(name) if name == &param.name)
        ) || matches!(
            &expr.kind,
            ExprKind::Call { callee, args, .. }
                if args.is_empty()
                    && matches!(
                        &callee.kind,
                        ExprKind::Member { object, field, .. }
                            if field == "sum"
                                && matches!(&object.kind, ExprKind::Ident(name) if name == &param.name)
                    )
        )
    };
    match body {
        LambdaBody::Expr(expr) => is_sum(expr),
        LambdaBody::Block(stmts) if stmts.len() == 1 => match &stmts[0].kind {
            StmtKind::Return(Some(expr)) | StmtKind::Expr(expr) => is_sum(expr),
            _ => false },
        _ => false }
}

fn kotlin_lambda_join_to_string(lambda: &Expression) -> Option<Expression> {
    let ExprKind::Lambda { params, body, .. } = &lambda.kind else {
        return None;
    };
    let param = params.first()?;
    let join_separator = |expr: &Expression| match &expr.kind {
        ExprKind::Call { callee, args, .. }
            if matches!(&callee.kind, ExprKind::Ident(name) if name == "__coll_join")
                && matches!(
                    args.first().map(|arg| &arg.value.kind),
                    Some(ExprKind::Ident(name)) if name == &param.name
                ) =>
        {
            args.get(1).map(|arg| arg.value.clone()).or_else(|| {
                Some(Expression::new(ExprKind::Lit(Literal::Str(
                    ", ".to_string(),
                ))))
            })
        }
        ExprKind::Call { callee, args, .. }
            if matches!(
                &callee.kind,
                ExprKind::Member { object, field, .. }
                    if field == "joinToString"
                        && matches!(&object.kind, ExprKind::Ident(name) if name == &param.name)
            ) =>
        {
            args.first().map(|arg| arg.value.clone()).or_else(|| {
                Some(Expression::new(ExprKind::Lit(Literal::Str(
                    ", ".to_string(),
                ))))
            })
        }
        _ => None };
    match body {
        LambdaBody::Expr(expr) => join_separator(expr),
        LambdaBody::Block(stmts) if stmts.len() == 1 => match &stmts[0].kind {
            StmtKind::Return(Some(expr)) | StmtKind::Expr(expr) => join_separator(expr),
            _ => None },
        _ => None }
}

fn kotlin_static_array_map_rewrite(expr: &Expression) -> Option<Expression> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if !matches!(&callee.kind, ExprKind::Ident(name) if name == "__array_map") || args.len() != 2 {
        return None;
    }
    let items = kotlin_static_array_items(&args[0].value)?;
    kotlin_apply_static_join_transform(&items, &args[1].value)
        .or_else(|| kotlin_apply_static_interpolation_transform(&items, &args[1].value))
}

fn kotlin_apply_static_join_transform(
    items: &[Expression],
    lambda: &Expression,
) -> Option<Expression> {
    let separator = kotlin_lambda_join_to_string(lambda)?;
    let mapped = items
        .iter()
        .map(|item| {
            Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__coll_join")),
                args: vec![
                    Argument::positional(item.clone()),
                    Argument::positional(separator.clone()),
                ],
                optional: false })
        })
        .collect();
    Some(kotlin_array_expr(mapped))
}

fn kotlin_apply_static_interpolation_transform(
    items: &[Expression],
    lambda: &Expression,
) -> Option<Expression> {
    let ExprKind::Lambda { params, body, .. } = &lambda.kind else {
        return None;
    };
    let param = params.first()?;
    let expr = match body {
        LambdaBody::Expr(expr) => expr.as_ref(),
        LambdaBody::Block(stmts) if stmts.len() == 1 => match &stmts[0].kind {
            StmtKind::Return(Some(expr)) | StmtKind::Expr(expr) => expr,
            _ => return None },
        _ => return None };
    let ExprKind::Interpolation(parts) = &expr.kind else {
        return None;
    };
    let mut mapped = Vec::new();
    for item in items {
        let mut text = String::new();
        for part in parts {
            match part {
                InterpolPart::Text(part) => text.push_str(part),
                InterpolPart::Expr(expr) => {
                    text.push_str(&kotlin_static_interp_expr(expr, &param.name, item)?);
                }
                InterpolPart::Formatted(expr, _) => {
                    text.push_str(&kotlin_static_interp_expr(expr, &param.name, item)?);
                }
            }
        }
        mapped.push(Expression::new(ExprKind::Lit(Literal::Str(text))));
    }
    Some(kotlin_array_expr(mapped))
}

fn kotlin_static_interp_expr(expr: &Expression, param: &str, item: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) if name == param => kotlin_static_value_to_string(item),
        ExprKind::Member { object, field, .. } if matches!(&object.kind, ExprKind::Ident(name) if name == param) =>
        {
            let ExprKind::Tuple(values) = &item.kind else {
                return None;
            };
            match field.as_str() {
                "first" => values.first().and_then(kotlin_static_value_to_string),
                "second" => values.get(1).and_then(kotlin_static_value_to_string),
                "third" => values.get(2).and_then(kotlin_static_value_to_string),
                _ => None }
        }
        ExprKind::Index { object, index, .. } if matches!(&object.kind, ExprKind::Ident(name) if name == param) =>
        {
            let ExprKind::Tuple(values) = &item.kind else {
                return None;
            };
            let ExprKind::Lit(Literal::Int(index)) = index.kind else {
                return None;
            };
            values
                .get(index as usize)
                .and_then(kotlin_static_value_to_string)
        }
        _ => None }
}

fn kotlin_static_value_to_string(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(value)) => Some(value.to_string()),
        ExprKind::Lit(Literal::Str(value)) => Some(value.clone()),
        ExprKind::Lit(Literal::Bool(value)) => Some(value.to_string()),
        _ => None }
}

fn kotlin_sum_static_items(items: &[Expression]) -> Option<Expression> {
    let mut sum = 0;
    for item in items {
        let ExprKind::Lit(Literal::Int(value)) = item.kind else {
            return None;
        };
        sum += value;
    }
    Some(Expression::int(sum))
}

fn kotlin_eval_sequence_lambda(
    body: &LambdaBody,
    param: &str,
    current: &KotlinSeqValue,
) -> Option<KotlinSeqValue> {
    match body {
        LambdaBody::Block(stmts) => kotlin_eval_sequence_stmts(stmts, param, current),
        LambdaBody::Expr(expr) => kotlin_eval_sequence_expr(expr, param, current) }
}

fn kotlin_eval_sequence_stmts(
    stmts: &[Statement],
    param: &str,
    current: &KotlinSeqValue,
) -> Option<KotlinSeqValue> {
    if stmts.len() != 1 {
        return None;
    }
    match &stmts[0].kind {
        StmtKind::Return(Some(expr)) | StmtKind::Expr(expr) => {
            kotlin_eval_sequence_expr(expr, param, current)
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body } => {
            if kotlin_eval_sequence_bool(cond, param, current)? {
                kotlin_eval_sequence_stmts(then_body, param, current)
            } else {
                for (elif_cond, elif_body) in elifs {
                    if kotlin_eval_sequence_bool(elif_cond, param, current)? {
                        return kotlin_eval_sequence_stmts(elif_body, param, current);
                    }
                }
                kotlin_eval_sequence_stmts(else_body.as_deref()?, param, current)
            }
        }
        _ => None }
}

fn kotlin_eval_sequence_bool_from_value(
    body: &LambdaBody,
    param: &str,
    current: &KotlinSeqValue,
) -> Option<bool> {
    match body {
        LambdaBody::Block(stmts) if stmts.len() == 1 => match &stmts[0].kind {
            StmtKind::Return(Some(expr)) | StmtKind::Expr(expr) => {
                kotlin_eval_sequence_bool(expr, param, current)
            }
            _ => None },
        LambdaBody::Expr(expr) => kotlin_eval_sequence_bool(expr, param, current),
        _ => None }
}

fn kotlin_seq_literal_value(expr: &Expression) -> Option<KotlinSeqValue> {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(value)) => Some(KotlinSeqValue::Int(*value)),
        ExprKind::Lit(Literal::Str(value)) => Some(KotlinSeqValue::Str(value.clone())),
        ExprKind::Lit(Literal::Null) => Some(KotlinSeqValue::Null),
        _ => None }
}

fn kotlin_seq_value_expr(value: &KotlinSeqValue) -> Option<Expression> {
    match value {
        KotlinSeqValue::Int(value) => Some(Expression::int(*value)),
        KotlinSeqValue::Str(value) => {
            Some(Expression::new(ExprKind::Lit(Literal::Str(value.clone()))))
        }
        KotlinSeqValue::Null => None }
}

fn kotlin_eval_sequence_expr(
    expr: &Expression,
    param: &str,
    current: &KotlinSeqValue,
) -> Option<KotlinSeqValue> {
    match &expr.kind {
        ExprKind::Lit(Literal::Null) => Some(KotlinSeqValue::Null),
        ExprKind::Lit(Literal::Int(value)) => Some(KotlinSeqValue::Int(*value)),
        ExprKind::Lit(Literal::Str(value)) => Some(KotlinSeqValue::Str(value.clone())),
        ExprKind::Ident(name) if name == param => Some(current.clone()),
        ExprKind::Call { callee, args, .. } if args.len() == 1 => {
            let ExprKind::Ident(name) = &callee.kind else {
                return None;
            };
            let arg_value = kotlin_eval_sequence_expr(&args[0].value, param, current)?;
            KOTLIN_SIMPLE_FUNCTIONS.with(|functions| {
                let functions = functions.borrow();
                let (params, body) = functions.get(name)?;
                if params.len() != 1 {
                    return None;
                }
                kotlin_eval_sequence_stmts(body, &params[0], &arg_value)
            })
        }
        ExprKind::Ternary { cond, then, else_ } => {
            if kotlin_eval_sequence_bool(cond, param, current)? {
                kotlin_eval_sequence_expr(then, param, current)
            } else {
                kotlin_eval_sequence_expr(else_, param, current)
            }
        }
        ExprKind::Binary { op, left, right } => {
            let l = kotlin_eval_sequence_expr(left, param, current)?;
            let r = kotlin_eval_sequence_expr(right, param, current)?;
            match (op, l, r) {
                (BinOp::Add, KotlinSeqValue::Int(l), KotlinSeqValue::Int(r)) => {
                    Some(KotlinSeqValue::Int(l + r))
                }
                (BinOp::Sub, KotlinSeqValue::Int(l), KotlinSeqValue::Int(r)) => {
                    Some(KotlinSeqValue::Int(l - r))
                }
                (BinOp::Mul, KotlinSeqValue::Int(l), KotlinSeqValue::Int(r)) => {
                    Some(KotlinSeqValue::Int(l * r))
                }
                (BinOp::Div, KotlinSeqValue::Int(l), KotlinSeqValue::Int(r)) if r != 0 => {
                    Some(KotlinSeqValue::Int(l / r))
                }
                (BinOp::Mod, KotlinSeqValue::Int(l), KotlinSeqValue::Int(r)) if r != 0 => {
                    Some(KotlinSeqValue::Int(l % r))
                }
                (BinOp::Add | BinOp::Concat, KotlinSeqValue::Str(l), KotlinSeqValue::Str(r)) => {
                    Some(KotlinSeqValue::Str(format!("{l}{r}")))
                }
                _ => None }
        }
        _ => None }
}

fn kotlin_eval_sequence_bool(
    expr: &Expression,
    param: &str,
    current: &KotlinSeqValue,
) -> Option<bool> {
    let ExprKind::Binary { op, left, right } = &expr.kind else {
        return None;
    };
    let l = kotlin_eval_sequence_comparable(left, param, current)?;
    let r = kotlin_eval_sequence_comparable(right, param, current)?;
    match (op, l, r) {
        (BinOp::Lt, KotlinSeqValue::Int(l), KotlinSeqValue::Int(r)) => Some(l < r),
        (BinOp::LtEq, KotlinSeqValue::Int(l), KotlinSeqValue::Int(r)) => Some(l <= r),
        (BinOp::Gt, KotlinSeqValue::Int(l), KotlinSeqValue::Int(r)) => Some(l > r),
        (BinOp::GtEq, KotlinSeqValue::Int(l), KotlinSeqValue::Int(r)) => Some(l >= r),
        (BinOp::Eq, KotlinSeqValue::Int(l), KotlinSeqValue::Int(r)) => Some(l == r),
        (BinOp::NotEq, KotlinSeqValue::Int(l), KotlinSeqValue::Int(r)) => Some(l != r),
        (BinOp::Eq, KotlinSeqValue::Str(l), KotlinSeqValue::Str(r)) => Some(l == r),
        (BinOp::NotEq, KotlinSeqValue::Str(l), KotlinSeqValue::Str(r)) => Some(l != r),
        _ => None }
}

fn kotlin_eval_sequence_comparable(
    expr: &Expression,
    param: &str,
    current: &KotlinSeqValue,
) -> Option<KotlinSeqValue> {
    match &expr.kind {
        ExprKind::Call { callee, args, .. }
            if matches!(&callee.kind, ExprKind::Ident(name) if name == "__coll_length")
                && args.len() == 1 =>
        {
            let value = kotlin_eval_sequence_expr(&args[0].value, param, current)?;
            match value {
                KotlinSeqValue::Str(value) => {
                    Some(KotlinSeqValue::Int(value.chars().count() as i64))
                }
                _ => None }
        }
        ExprKind::Member { object, field, .. } if field == "length" => {
            let value = kotlin_eval_sequence_expr(object, param, current)?;
            match value {
                KotlinSeqValue::Str(value) => {
                    Some(KotlinSeqValue::Int(value.chars().count() as i64))
                }
                _ => None }
        }
        _ => kotlin_eval_sequence_expr(expr, param, current) }
}

#[derive(Debug, Clone, Default)]
struct KotlinOperatorInfo {
    returns: HashMap<String, Option<String>> }

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

/// `val v = try { … } finally { … }` was an IIFE, and the wrapper lambda's
/// captured reads went STALE (a local fn's write to `executed` landed in
/// the enclosing frame while the `finally` read the lambda's copy). In
/// DECLARATION position no lambda is needed: run the try as a STATEMENT
/// assigning a temp.
fn kotlin_unlift_try_inits(stmts: &mut Vec<Statement>) {
    fn returns_to_assign(stmts: &mut [Statement], tmp: &str) {
        for stmt in stmts {
            match &mut stmt.kind {
                StmtKind::Return(Some(e)) => {
                    let value = e.clone();
                    stmt.kind = StmtKind::Expr(Expression::new(ExprKind::Assign {
                        target: Box::new(Expression::ident(tmp)),
                        value: Box::new(value) }));
                }
                StmtKind::Try { body, catches, else_body, finally } => {
                    returns_to_assign(body, tmp);
                    for c in catches {
                        returns_to_assign(&mut c.body, tmp);
                    }
                    if let Some(b) = else_body {
                        returns_to_assign(b, tmp);
                    }
                    if let Some(b) = finally {
                        returns_to_assign(b, tmp);
                    }
                }
                StmtKind::If { then_body, elifs, else_body, .. } => {
                    returns_to_assign(then_body, tmp);
                    for (_, b) in elifs {
                        returns_to_assign(b, tmp);
                    }
                    if let Some(b) = else_body {
                        returns_to_assign(b, tmp);
                    }
                }
                StmtKind::Block(body) => returns_to_assign(body, tmp),
                _ => {}
            }
        }
    }
    // Recurse into every nested statement body (and lambda bodies through
    // the expression walker) FIRST, then splice at this level. A Block
    // wrapper is NOT an option: the shared emitter opens a lexical scope
    // for blocks that declare `val`s, which would hide the binding.
    for stmt in stmts.iter_mut() {
        match &mut stmt.kind {
            StmtKind::FunctionDecl { body, .. } => kotlin_unlift_try_inits(body),
            StmtKind::If { then_body, elifs, else_body, .. } => {
                kotlin_unlift_try_inits(then_body);
                for (_, b) in elifs {
                    kotlin_unlift_try_inits(b);
                }
                if let Some(b) = else_body {
                    kotlin_unlift_try_inits(b);
                }
            }
            StmtKind::While { body, .. } | StmtKind::ForIn { body, .. } => {
                kotlin_unlift_try_inits(body);
            }
            StmtKind::Try { body, catches, else_body, finally } => {
                kotlin_unlift_try_inits(body);
                for c in catches {
                    kotlin_unlift_try_inits(&mut c.body);
                }
                if let Some(b) = else_body {
                    kotlin_unlift_try_inits(b);
                }
                if let Some(b) = finally {
                    kotlin_unlift_try_inits(b);
                }
            }
            StmtKind::Block(body) => kotlin_unlift_try_inits(body),
            StmtKind::ClassDecl { members, .. } => {
                for member in members.iter_mut() {
                    if let ClassMember::Method(m) = member
                        && let StmtKind::FunctionDecl { body, .. } = &mut m.kind
                    {
                        kotlin_unlift_try_inits(body);
                    }
                }
            }
            _ => {}
        }
        stmt.walk_exprs_mut(&mut |e| {
            if let ExprKind::Lambda { body: LambdaBody::Block(b), .. } = &mut e.kind {
                kotlin_unlift_try_inits(b);
            }
        });
    }

    let mut i = 0;
    while i < stmts.len() {
        let matched = {
            match &stmts[i].kind {
                StmtKind::VarDecl { declarations, .. } if declarations.len() == 1 => {
                    match &declarations[0].init {
                        Some(init) => match &init.kind {
                            ExprKind::Call { callee, args, .. } if args.is_empty() => {
                                matches!(&callee.kind,
                                    ExprKind::Lambda { params, body: LambdaBody::Block(lbody), .. }
                                        if params.is_empty()
                                            && lbody.len() == 1
                                            && matches!(lbody[0].kind, StmtKind::Try { .. }))
                            }
                            _ => false },
                        None => false }
                }
                _ => false }
        };
        if !matched {
            i += 1;
            continue;
        }
        let StmtKind::VarDecl { declarations, kind } = &stmts[i].kind else {
            unreachable!()
        };
        let ExprKind::Call { callee, .. } = &declarations[0].init.as_ref().unwrap().kind
        else {
            unreachable!()
        };
        let ExprKind::Lambda { body: LambdaBody::Block(lbody), .. } = &callee.kind else {
            unreachable!()
        };
        let tmp = format!("__kt_try_{i}");
        let mut try_stmt = lbody[0].clone();
        returns_to_assign(std::slice::from_mut(&mut try_stmt), &tmp);
        let mut new_decl = declarations[0].clone();
        new_decl.init = Some(Expression::ident(&tmp));
        let decl_kind = kind.clone();
        let tmp_decl = Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(tmp.clone()),
                type_hint: None,
                init: Some(Expression::null()),
                array_bounds: None,
                with_events: false }],
            kind: VarDeclKind::Var });
        stmts[i] = tmp_decl;
        stmts.insert(i + 1, try_stmt);
        stmts.insert(
            i + 2,
            Statement::new(StmtKind::VarDecl {
                declarations: vec![new_decl],
                kind: decl_kind }),
        );
        i += 3;
    }
}

fn kotlin_local_capture_params(names: &[String]) -> Vec<Param> {
    names
        .iter()
        .map(|name| Param {
            name: name.clone(),
            type_hint: None,
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false })
        .collect()
}

fn collect_kotlin_free_locals_in_stmts(
    stmts: &[Statement],
    locals: &KotlinLocalTypes,
    bound: &mut HashSet<String>,
    out: &mut HashSet<String>,
) {
    for stmt in stmts {
        match &stmt.kind {
            StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
                collect_kotlin_free_locals_in_expr(expr, locals, bound, out);
            }
            StmtKind::Throw { expr, cause } => {
                if let Some(expr) = expr {
                    collect_kotlin_free_locals_in_expr(expr, locals, bound, out);
                }
                if let Some(cause) = cause {
                    collect_kotlin_free_locals_in_expr(cause, locals, bound, out);
                }
            }
            StmtKind::Echo(exprs) => {
                for expr in exprs {
                    collect_kotlin_free_locals_in_expr(expr, locals, bound, out);
                }
            }
            StmtKind::VarDecl { declarations, .. } => {
                for decl in declarations {
                    if let Some(init) = &decl.init {
                        collect_kotlin_free_locals_in_expr(init, locals, bound, out);
                    }
                    collect_binding_names(&decl.pattern, bound);
                }
            }
            StmtKind::Block(body) | StmtKind::NamespaceDecl { body, .. } => {
                let mut inner_bound = bound.clone();
                collect_kotlin_free_locals_in_stmts(body, locals, &mut inner_bound, out);
            }
            StmtKind::If {
                cond,
                then_body,
                elifs,
                else_body } => {
                collect_kotlin_free_locals_in_expr(cond, locals, bound, out);
                collect_kotlin_free_locals_in_stmts(then_body, locals, &mut bound.clone(), out);
                for (elif_cond, elif_body) in elifs {
                    collect_kotlin_free_locals_in_expr(elif_cond, locals, bound, out);
                    collect_kotlin_free_locals_in_stmts(elif_body, locals, &mut bound.clone(), out);
                }
                if let Some(else_body) = else_body {
                    collect_kotlin_free_locals_in_stmts(else_body, locals, &mut bound.clone(), out);
                }
            }
            StmtKind::While {
                cond,
                body,
                else_body } => {
                collect_kotlin_free_locals_in_expr(cond, locals, bound, out);
                collect_kotlin_free_locals_in_stmts(body, locals, &mut bound.clone(), out);
                if let Some(else_body) = else_body {
                    collect_kotlin_free_locals_in_stmts(else_body, locals, &mut bound.clone(), out);
                }
            }
            StmtKind::For {
                init,
                cond,
                update,
                body } => {
                let mut inner_bound = bound.clone();
                if let Some(init) = init {
                    collect_kotlin_free_locals_in_stmts(
                        std::slice::from_ref(init),
                        locals,
                        &mut inner_bound,
                        out,
                    );
                }
                if let Some(cond) = cond {
                    collect_kotlin_free_locals_in_expr(cond, locals, &inner_bound, out);
                }
                if let Some(update) = update {
                    collect_kotlin_free_locals_in_expr(update, locals, &inner_bound, out);
                }
                collect_kotlin_free_locals_in_stmts(body, locals, &mut inner_bound, out);
            }
            StmtKind::ForIn {
                var,
                key,
                iter,
                body,
                else_body,
                ..
            } => {
                collect_kotlin_free_locals_in_expr(iter, locals, bound, out);
                let mut inner_bound = bound.clone();
                inner_bound.insert(var.clone());
                if let Some(key) = key {
                    inner_bound.insert(key.clone());
                }
                collect_kotlin_free_locals_in_stmts(body, locals, &mut inner_bound, out);
                if let Some(else_body) = else_body {
                    collect_kotlin_free_locals_in_stmts(else_body, locals, &mut bound.clone(), out);
                }
            }
            StmtKind::Try {
                body,
                catches,
                else_body,
                finally } => {
                collect_kotlin_free_locals_in_stmts(body, locals, &mut bound.clone(), out);
                for catch in catches {
                    let mut catch_bound = bound.clone();
                    if let Some(name) = &catch.var_name {
                        catch_bound.insert(name.clone());
                    }
                    if let Some(name) = &catch.stack_var {
                        catch_bound.insert(name.clone());
                    }
                    if let Some(when_clause) = &catch.when_clause {
                        collect_kotlin_free_locals_in_expr(when_clause, locals, &catch_bound, out);
                    }
                    collect_kotlin_free_locals_in_stmts(&catch.body, locals, &mut catch_bound, out);
                }
                if let Some(else_body) = else_body {
                    collect_kotlin_free_locals_in_stmts(else_body, locals, &mut bound.clone(), out);
                }
                if let Some(finally) = finally {
                    collect_kotlin_free_locals_in_stmts(finally, locals, &mut bound.clone(), out);
                }
            }
            _ => {}
        }
    }
}

fn collect_kotlin_free_locals_in_expr(
    expr: &Expression,
    locals: &KotlinLocalTypes,
    bound: &HashSet<String>,
    out: &mut HashSet<String>,
) {
    match &expr.kind {
        ExprKind::Ident(name) => {
            if locals.contains_key(name) && !bound.contains(name) {
                out.insert(name.clone());
            }
        }
        ExprKind::Binary { left, right, .. } => {
            collect_kotlin_free_locals_in_expr(left, locals, bound, out);
            collect_kotlin_free_locals_in_expr(right, locals, bound, out);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::Await(expr)
        | ExprKind::Yield(Some(expr))
        | ExprKind::Delete(expr) => collect_kotlin_free_locals_in_expr(expr, locals, bound, out),
        ExprKind::Assign { target, value } => {
            collect_kotlin_free_locals_in_expr(target, locals, bound, out);
            collect_kotlin_free_locals_in_expr(value, locals, bound, out);
        }
        ExprKind::Call { callee, args, .. } => {
            collect_kotlin_free_locals_in_expr(callee, locals, bound, out);
            for arg in args {
                collect_kotlin_free_locals_in_expr(&arg.value, locals, bound, out);
            }
        }
        ExprKind::New { class, args } => {
            collect_kotlin_free_locals_in_expr(class, locals, bound, out);
            for arg in args {
                collect_kotlin_free_locals_in_expr(&arg.value, locals, bound, out);
            }
        }
        ExprKind::Member { object, .. } => {
            collect_kotlin_free_locals_in_expr(object, locals, bound, out);
        }
        ExprKind::Index { object, index, .. } => {
            collect_kotlin_free_locals_in_expr(object, locals, bound, out);
            collect_kotlin_free_locals_in_expr(index, locals, bound, out);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            collect_kotlin_free_locals_in_expr(cond, locals, bound, out);
            collect_kotlin_free_locals_in_expr(then, locals, bound, out);
            collect_kotlin_free_locals_in_expr(else_, locals, bound, out);
        }
        ExprKind::NullCoalesce { left, right } => {
            collect_kotlin_free_locals_in_expr(left, locals, bound, out);
            collect_kotlin_free_locals_in_expr(right, locals, bound, out);
        }
        ExprKind::Array(elems) => {
            for elem in elems {
                if let Some(key) = &elem.key {
                    collect_kotlin_free_locals_in_expr(key, locals, bound, out);
                }
                collect_kotlin_free_locals_in_expr(&elem.value, locals, bound, out);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                if let ObjectProperty::KeyValue { key, value } = prop {
                    collect_kotlin_free_locals_in_expr(key, locals, bound, out);
                    collect_kotlin_free_locals_in_expr(value, locals, bound, out);
                }
            }
        }
        ExprKind::Tuple(items) | ExprKind::Sequence(items) => {
            for item in items {
                collect_kotlin_free_locals_in_expr(item, locals, bound, out);
            }
        }
        ExprKind::Range { start, end, .. } => {
            collect_kotlin_free_locals_in_expr(start, locals, bound, out);
            collect_kotlin_free_locals_in_expr(end, locals, bound, out);
        }
        ExprKind::Lambda { params, body, .. } => {
            let mut lambda_bound = bound.clone();
            for param in params {
                lambda_bound.insert(param.name.clone());
            }
            match body {
                LambdaBody::Expr(expr) => {
                    collect_kotlin_free_locals_in_expr(expr, locals, &lambda_bound, out)
                }
                LambdaBody::Block(stmts) => {
                    collect_kotlin_free_locals_in_stmts(stmts, locals, &mut lambda_bound, out)
                }
            }
        }
        _ => {}
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
                // Type from the PRE-normalize init: rewrites erase the
                // spelling (`FloatArray(2)` becomes `__kt_array_init(2)`),
                // and the element type drives Double rendering.
                let pre_normalize_ty = decl
                    .init
                    .as_ref()
                    .and_then(|e| kotlin_expr_type(e, locals, operators));
                if let Some(init) = &mut decl.init {
                    normalize_kotlin_operator_expr(init, operators, locals);
                    let original_init = init.clone();
                    if let Some(materialized) = kotlin_materialize_generate_sequence(init, None) {
                        *init = materialized;
                        if let BindingPattern::Ident(name) = &decl.pattern {
                            KOTLIN_STATIC_VALUES.with(|map| {
                                map.borrow_mut().insert(name.clone(), init.clone());
                            });
                        }
                    } else if let BindingPattern::Ident(name) = &decl.pattern {
                        if kotlin_is_sequence_source(&original_init) {
                            KOTLIN_SEQUENCE_SOURCES.with(|map| {
                                map.borrow_mut().insert(name.clone(), original_init);
                            });
                        }
                        if let Some(collection_ty) = kotlin_literal_keyed_collection_type(init) {
                            KOTLIN_KEYED_COLLECTION_TYPES.with(|map| {
                                map.borrow_mut()
                                    .insert(name.clone(), collection_ty.to_string());
                            });
                        }
                        if matches!(init.kind, ExprKind::Array(_)) {
                            KOTLIN_STATIC_VALUES.with(|map| {
                                map.borrow_mut().insert(name.clone(), init.clone());
                            });
                        }
                    }
                }
                if let BindingPattern::Ident(name) = &decl.pattern {
                    let inferred = decl
                        .type_hint
                        .as_deref()
                        .map(str::to_string)
                        .or_else(|| pre_normalize_ty.clone())
                        .or_else(|| {
                            decl.init
                                .as_ref()
                                .and_then(|e| kotlin_expr_type(e, locals, operators))
                        });
                    if let Some(ty) = inferred {
                        if decl.type_hint.is_none() {
                            decl.type_hint = Some(ty.clone().into());
                        }
                        locals.insert(name.clone(), ty.to_string());
                    }
                }
                // A receiver-typed function VALUE (`val upper: String.() ->
                // String = { uppercase() }`): the lambda gains a leading
                // `__recv` parameter and its body binds `this`/bare members
                // to it — the call-site rewrite passes the receiver first.
                if let Some(hint) = decl.type_hint.as_deref().map(str::to_string)
                    && kotlin_hint_is_receiver_fn(&hint)
                    && let Some(init) = &mut decl.init
                    && let ExprKind::Lambda { params, body, .. } = &mut init.kind
                    && params.first().map(|p| p.name.as_str()) != Some("__recv")
                {
                    let mut skip = HashSet::new();
                    for p in params.iter() {
                        skip.insert(p.name.clone());
                    }
                    match body {
                        LambdaBody::Block(stmts) => {
                            kotlin_local_binding_names(stmts, &mut skip);
                            for stmt in stmts.iter_mut() {
                                kotlin_rewrite_receiver_refs_stmt(stmt, "__recv", &skip);
                            }
                        }
                        LambdaBody::Expr(e) => {
                            kotlin_rewrite_receiver_refs_expr(e, "__recv", &skip);
                        }
                    }
                    params.insert(
                        0,
                        kotlin_local_capture_params(&["__recv".to_string()])
                            .pop()
                            .unwrap(),
                    );
                }
            }
        }
        StmtKind::FunctionDecl { params, body, .. } => {
            let mut fn_locals = locals.clone();
            for param in params.iter_mut() {
                // DEFAULT expressions normalize too — the typed-array
                // factories materialize in this pass now, so an unvisited
                // `rest: IntArray = intArrayOf(1, 2)` stayed an unresolved
                // call and the default was undefined.
                if let Some(default) = &mut param.default {
                    normalize_kotlin_operator_expr(default, operators, locals);
                }
                if let Some(ty) = &param.type_hint {
                    fn_locals.insert(param.name.clone(), ty.clone().to_string());
                }
            }
            normalize_kotlin_operator_stmts(body, operators, &mut fn_locals);
        }
        StmtKind::ClassDecl { name: class_name, members, .. } => {
            let class_props: HashSet<String> = CLASS_PROPERTIES.with(|m| {
                m.borrow()
                    .get(class_name.as_str())
                    .map(|props| props.keys().cloned().collect())
                    .unwrap_or_default()
            });
            for member in members {
                match member {
                    ClassMember::Field {
                        init: Some(init), ..
                    } => normalize_kotlin_operator_expr(init, operators, locals),
                    ClassMember::Method(method) => {
                        // A bare CALL of a function-valued PROPERTY
                        // (`op(v)` where `op` is `val op: (Int) -> Int`)
                        // never went through implicit-self (that covers
                        // reads) — bind the callee to `this`. Locals and
                        // params shadow.
                        if let StmtKind::FunctionDecl { params, body, .. } = &mut method.kind {
                            let mut shadowed: HashSet<String> = params
                                .iter()
                                .map(|p| p.name.clone())
                                .collect();
                            kotlin_local_binding_names(body, &mut shadowed);
                            for stmt in body.iter_mut() {
                                stmt.walk_exprs_mut(&mut |e| {
                                    if let ExprKind::Call { callee, .. } = &mut e.kind
                                        && let ExprKind::Ident(n) = &callee.kind
                                        && class_props.contains(n.as_str())
                                        && !shadowed.contains(n.as_str())
                                        && !KOTLIN_TOP_FN_PARAMS
                                            .with(|m| m.borrow().contains_key(n.as_str()))
                                    {
                                        let field = n.clone();
                                        **callee = Expression::new(ExprKind::Member {
                                            object: Box::new(Expression::new(
                                                ExprKind::This,
                                            )),
                                            field,
                                            null_safe: false });
                                    }
                                });
                            }
                        }
                        normalize_kotlin_operator_stmt(method, operators, locals);
                    }
                    ClassMember::Constructor { params, body, .. } => {
                        let mut ctor_locals = locals.clone();
                        for param in params {
                            if let Some(ty) = &param.type_hint {
                                ctor_locals.insert(param.name.clone(), ty.clone().to_string());
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
            else_body } => {
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
            else_body } => {
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
            body } => {
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
            var,
            key,
            iter,
            body,
            else_body,
            ..
        } => {
            normalize_kotlin_operator_expr(iter, operators, locals);
            if kotlin_expr_type(iter, locals, operators)
                .as_deref()
                .is_some_and(kotlin_type_is_set_like)
            {
                *iter = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__kt_to_list")),
                    args: vec![Argument::positional(iter.clone())],
                    optional: false });
            }
            // A loop variable that reuses an enclosing name is a NEW
            // binding in Kotlin — renaming keeps the outer var untouched
            // (`for (outer in …)` was mutating `outer`).
            if locals.contains_key(var.as_str()) {
                let renamed = format!("{var}__shadow");
                let old = var.clone();
                for s in body.iter_mut() {
                    s.walk_exprs_mut(&mut |e| {
                        if let ExprKind::Ident(n) = &e.kind
                            && *n == old
                        {
                            *e = Expression::ident(&renamed);
                        }
                    });
                }
                *var = renamed;
            }
            let mut for_locals = locals.clone();
            for_locals.insert(var.clone(), "Any".to_string());
            if let Some(key) = key {
                for_locals.insert(key.clone(), "Any".to_string());
            }
            normalize_kotlin_operator_stmts(body, operators, &mut for_locals);
            if let Some(else_body) = else_body {
                normalize_kotlin_operator_stmts(else_body, operators, &mut locals.clone());
            }
        }
        // Statement `when` — its conditions and arm bodies carry full
        // expressions; skipping them left `in 2..3` (and every other
        // rewrite) unlowered inside statement-position `when`s.
        StmtKind::Switch { expr, cases, default } => {
            normalize_kotlin_operator_expr(expr, operators, locals);
            for case in cases {
                for cond in &mut case.conditions {
                    match cond {
                        CaseCondition::Value(e) => {
                            normalize_kotlin_operator_expr(e, operators, locals)
                        }
                        CaseCondition::Range { from, to } => {
                            normalize_kotlin_operator_expr(from, operators, locals);
                            normalize_kotlin_operator_expr(to, operators, locals);
                        }
                        CaseCondition::Comparison { expr, .. } => {
                            normalize_kotlin_operator_expr(expr, operators, locals)
                        }
                    }
                }
                normalize_kotlin_operator_stmts(&mut case.body, operators, &mut locals.clone());
            }
            if let Some(default) = default {
                normalize_kotlin_operator_stmts(default, operators, &mut locals.clone());
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally } => {
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
        ExprKind::Binary { op, left, right } => {
            normalize_kotlin_operator_expr(left, operators, locals);
            normalize_kotlin_operator_expr(right, operators, locals);
            // Kotlin `Map + map/pair` and `Map - key` are `plus`/`minus`
            // operators, not numeric arithmetic. The shared BinOp path would
            // coerce both sides to numbers (`NaN`).
            if matches!(op, BinOp::In)
                && kotlin_expr_type(right, locals, operators)
                    .as_deref()
                    .map(|t| {
                        t.split('<').next().unwrap_or(t).rsplit('.').next().unwrap_or(t)
                            == "String"
                    })
                    .unwrap_or(false)
            {
                // Kotlin's `in` on a String is `contains`, not the property
                // probe `ecma_in_operator` routes to.
                *expr = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__kt_contains_ic")),
                    args: vec![
                        Argument::positional((**right).clone()),
                        Argument::positional((**left).clone()),
                        Argument::positional(Expression::new(ExprKind::Lit(Literal::Bool(
                            false,
                        )))),
                    ],
                    optional: false });
                return;
            }
            // A Double in STRING concatenation renders Kotlin-style:
            // `"red:" + 1.0` is `red:1.0`, never `red:1`.
            if *op == BinOp::Add {
                let left_str = matches!(left.kind, ExprKind::Lit(Literal::Str(_)))
                    || kotlin_expr_type(left, locals, operators).as_deref() == Some("String");
                let right_str = matches!(right.kind, ExprKind::Lit(Literal::Str(_)))
                    || kotlin_expr_type(right, locals, operators).as_deref() == Some("String");
                let wrap = |e: &mut Box<Expression>| {
                    **e = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__kt_double_str")),
                        args: vec![Argument::positional((**e).clone())],
                        optional: false });
                };
                if left_str
                    && !right_str
                    && kotlin_expr_type(right, locals, operators)
                        .as_deref()
                        .is_some_and(kotlin_type_is_double_like)
                {
                    wrap(right);
                }
                if right_str
                    && !left_str
                    && kotlin_expr_type(left, locals, operators)
                        .as_deref()
                        .is_some_and(kotlin_type_is_double_like)
                {
                    wrap(left);
                }
            }
            if matches!(op, BinOp::Add | BinOp::Sub)
                && kotlin_expr_type(left, locals, operators)
                    .as_deref()
                    .is_some_and(|ty| {
                        kotlin_type_is_plain_map_like(ty) || kotlin_type_is_list_like(ty)
                    })
            {
                let target = if matches!(op, BinOp::Add) {
                    "__kt_plus"
                } else {
                    "__kt_minus"
                };
                *expr = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident(target)),
                    args: vec![
                        Argument::positional((**left).clone()),
                        Argument::positional((**right).clone()),
                    ],
                    optional: false });
            }
        }
        ExprKind::Unary { expr: inner, .. }
        | ExprKind::Cast { expr: inner, .. }
        | ExprKind::Await(inner)
        | ExprKind::Yield(Some(inner))
        | ExprKind::Delete(inner) => normalize_kotlin_operator_expr(inner, operators, locals),
        // `when` arms carry full EXPRESSIONS — skipping them here meant
        // `in 70..79` (and every other Kotlin rewrite) never lowered inside
        // a `when`, so range guards always missed.
        ExprKind::Match { subject, arms } => {
            normalize_kotlin_operator_expr(subject, operators, locals);
            for arm in arms {
                if let Some(conds) = &mut arm.conditions {
                    for c in conds {
                        normalize_kotlin_operator_expr(c, operators, locals);
                    }
                }
                normalize_kotlin_operator_expr(&mut arm.body, operators, locals);
            }
        }
        ExprKind::Index { object, index, .. } => {
            normalize_kotlin_operator_expr(object, operators, locals);
            normalize_kotlin_operator_expr(index, operators, locals);
            // `sb[i]` — a StringBuilder indexes its BUFFER, not its properties.
            if kotlin_expr_type(object, locals, operators).as_deref()
                == Some("StringBuilder")
            {
                *expr = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Member {
                        object: object.clone(),
                        field: "charAt".to_string(),
                        null_safe: false })),
                    args: vec![Argument::positional((**index).clone())],
                    optional: false });
            } else if kotlin_expr_type(object, locals, operators)
                .as_deref()
                .map(|t| t.split('<').next().unwrap_or(t) == "String")
                .unwrap_or(false)
            {
                // `s[i]` on a String THROWS out of bounds in Kotlin.
                *expr = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__kt_char_at")),
                    args: vec![
                        Argument::positional((**object).clone()),
                        Argument::positional((**index).clone()),
                    ],
                    optional: false });
            }
        }
        ExprKind::Index { object, index, .. } => {
            normalize_kotlin_operator_expr(object, operators, locals);
            normalize_kotlin_operator_expr(index, operators, locals);
            // `sb[i]` — a StringBuilder indexes its BUFFER, not its properties.
            if kotlin_expr_type(object, locals, operators).as_deref()
                == Some("StringBuilder")
            {
                *expr = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Member {
                        object: object.clone(),
                        field: "charAt".to_string(),
                        null_safe: false })),
                    args: vec![Argument::positional((**index).clone())],
                    optional: false });
            } else if kotlin_expr_type(object, locals, operators)
                .as_deref()
                .map(|t| t.split('<').next().unwrap_or(t) == "String")
                .unwrap_or(false)
            {
                // `s[i]` on a String THROWS out of bounds in Kotlin.
                *expr = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__kt_char_at")),
                    args: vec![
                        Argument::positional((**object).clone()),
                        Argument::positional((**index).clone()),
                    ],
                    optional: false });
            }
        }
        ExprKind::Member { object, .. } => {
            normalize_kotlin_operator_expr(object, operators, locals);
            // kotlin.time.Duration constants and whole-unit accessors — a
            // Duration IS its millisecond count.
            if let ExprKind::Member { object: obj2, field, .. } = &expr.kind {
                if matches!(&obj2.kind, ExprKind::Ident(n) if n == "Duration") {
                    match field.as_str() {
                        "ZERO" => {
                            *expr = Expression::new(ExprKind::Lit(Literal::Float(0.0)));
                            return;
                        }
                        "INFINITE" => {
                            *expr =
                                Expression::new(ExprKind::Lit(Literal::Float(f64::INFINITY)));
                            return;
                        }
                        _ => {}
                    }
                }
                if matches!(field.as_str(), "isSuccess" | "isFailure")
                    && kotlin_expr_type(obj2, locals, operators)
                        .as_deref()
                        .is_some_and(kotlin_type_is_result)
                {
                    let ok = kotlin_result_prop((**obj2).clone(), "__kt_ok");
                    *expr = if field == "isSuccess" {
                        ok
                    } else {
                        Expression::new(ExprKind::Unary {
                            op: UnaryOp::Not,
                            expr: Box::new(ok) })
                    };
                    return;
                }
                if let Some((num, den)) = kotlin_duration_whole_factor(field) {
                    let obj = (**obj2).clone();
                    *expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__kt_dur_whole")),
                        args: vec![
                            Argument::positional(obj),
                            Argument::positional(Expression::new(ExprKind::Lit(
                                Literal::Float(num),
                            ))),
                            Argument::positional(Expression::new(ExprKind::Lit(
                                Literal::Float(den),
                            ))),
                        ],
                        optional: false });
                    return;
                }
            }
            // `kotlin.math.PI` — the namespace tree stores lowercase-canonical
            // keys and Kotlin is case-sensitive, so uppercase leaves (the
            // constants) missed. The kotlin.* stdlib surface is case-stable,
            // so the whole path folds.
            if let Some(path) = dotted_expr_path_of(expr) {
                // The same shape C#'s walker uses for `System.Math.PI`: a
                // compile-time constant is a LITERAL, not a lookup —
                // `[namespace_constants]` only serves two-segment names.
                match path.as_str() {
                    "kotlin.math.PI" => {
                        *expr = Expression::new(ExprKind::Lit(Literal::Float(
                            std::f64::consts::PI,
                        )));
                        return;
                    }
                    "kotlin.math.E" => {
                        *expr = Expression::new(ExprKind::Lit(Literal::Float(
                            std::f64::consts::E,
                        )));
                        return;
                    }
                    _ => {}
                }
                if let Some(rest) = path.strip_prefix("kotlin.") {
                    if rest.chars().any(|c| c.is_ascii_uppercase()) {
                        let lowered = path.to_ascii_lowercase();
                        let mut segs = lowered.split('.');
                        let mut built = Expression::ident(segs.next().unwrap());
                        for seg in segs {
                            built = Expression::new(ExprKind::Member {
                                object: Box::new(built),
                                field: seg.to_string(),
                                null_safe: false });
                        }
                        *expr = built;
                    }
                }
                // A bare `java.*` constant read (`IsoFields.WEEK_BASED_YEAR`)
                // never reaches the tree through the call path — resolve the
                // Const leaf through the COMMON resolver, same data the jvm
                // platform registered.
                if path.starts_with("java.") {
                    let lowered = path.to_ascii_lowercase();
                    let mut segments: Vec<&str> = vec!["jvm"];
                    segments.extend(lowered.split('.'));
                    if let Some(
                        vybe_compiler::primitives::namespaces::ResolutionTarget::Const(value),
                    ) = vybe_compiler::primitives::namespaces::resolve_path(&segments)
                    {
                        use vybe_runtime::Value;
                        let lit = match value {
                            Value::String(s) => Some(Literal::Str(s.to_string().into())),
                            Value::F64(v) => Some(Literal::Float(v)),
                            Value::Bool(b) => Some(Literal::Bool(b)),
                            _ => None };
                        if let Some(lit) = lit {
                            *expr = Expression::new(ExprKind::Lit(lit));
                            return;
                        }
                    }
                }
            }
        }
        ExprKind::Assign { target, value } => {
            if let ExprKind::Index { object, index, .. } = &mut target.kind {
                normalize_kotlin_operator_expr(object, operators, locals);
                normalize_kotlin_operator_expr(index, operators, locals);
            } else {
                normalize_kotlin_operator_expr(target, operators, locals);
            }
            normalize_kotlin_operator_expr(value, operators, locals);
        }
        ExprKind::Call { callee, args, .. } => {
            normalize_kotlin_operator_expr(callee, operators, locals);
            for arg in &mut *args {
                normalize_kotlin_operator_expr(&mut arg.value, operators, locals);
            }
            if args.is_empty() {
                if let ExprKind::Lambda {
                    params,
                    body: LambdaBody::Block(body),
                    captures,
                    ..
                } = &mut callee.kind
                {
                    if params.is_empty() && captures.is_empty() {
                        let mut free = HashSet::new();
                        collect_kotlin_free_locals_in_stmts(
                            body,
                            locals,
                            &mut HashSet::new(),
                            &mut free,
                        );
                        if !free.is_empty() {
                            let mut names: Vec<String> = free.into_iter().collect();
                            names.sort();
                            *params = kotlin_local_capture_params(&names);
                            args.extend(
                                names
                                    .into_iter()
                                    .map(|name| Argument::positional(Expression::ident(&name))),
                            );
                        }
                    }
                }
            }
            if matches!(&callee.kind, ExprKind::Ident(name) if name == "error") {
                *expr = kotlin_error_throw_expr(args);
                return;
            }
            if let ExprKind::Member { object, field, .. } = &mut callee.kind
                && matches!(
                    field.as_str(),
                    "filter" | "filterNot" | "map" | "forEach" | "fold" | "reduce" | "any" | "all"
                )
                && let Some(source) = kotlin_generated_dict_items_source(object)
                && kotlin_expr_type(&source, locals, operators)
                    .as_deref()
                    .is_some_and(kotlin_type_is_set_like)
            {
                *object = Box::new(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__kt_to_list")),
                    args: vec![Argument::positional(source)],
                    optional: false }));
            }
            if let ExprKind::Ident(name) = &callee.kind
                && name == "__kt_class_of"
                && args.len() == 1
                && let ExprKind::Ident(value_name) = &args[0].value.kind
                && let Some(type_name) = locals.get(value_name)
            {
                *expr = kotlin_class_literal_expr(type_name);
                return;
            }
            if let ExprKind::Ident(name) = &callee.kind
                && is_user_class_name(name)
            {
                let normalized_args = kotlin_normalized_constructor_args(name, args);
                *expr = kotlin_user_constructor_call(name, &normalized_args).unwrap_or_else(|| {
                    Expression::new(ExprKind::New {
                        class: Box::new(Expression::ident(name)),
                        args: normalized_args })
                });
                return;
            }
            if args.is_empty()
                && let ExprKind::Member { object, field, .. } = &callee.kind
            {
                match field.as_str() {
                    "toSet" | "toMutableSet" => {
                        *expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident("__kt_to_set")),
                            args: vec![Argument::positional((**object).clone())],
                            optional: false });
                        return;
                    }
                    "toList" | "toMutableList" | "toTypedArray" => {
                        if kotlin_expr_type(object, locals, operators)
                            .as_deref()
                            .is_some_and(kotlin_type_is_set_like)
                        {
                            *expr = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::ident("__kt_to_list")),
                                args: vec![Argument::positional((**object).clone())],
                                optional: false });
                            return;
                        }
                    }
                    _ => {}
                }
            }
            if let ExprKind::Member { object, field, .. } = &callee.kind
                && kotlin_expr_type(object, locals, operators)
                    .as_deref()
                    .is_some_and(kotlin_type_is_set_like)
                && matches!(
                    field.as_str(),
                    "map"
                        | "filter"
                        | "fold"
                        | "reduce"
                        | "any"
                        | "all"
                        | "none"
                        | "forEach"
                        | "sum"
                        | "min"
                        | "minOrNull"
                        | "max"
                        | "maxOrNull"
                        | "joinToString"
                        | "take"
                        | "drop"
                        | "first"
                        | "firstOrNull"
                        | "last"
                        | "lastOrNull"
                        | "elementAt"
                        | "count"
                        | "average"
                        | "toSortedSet"
                        | "iterator"
                        | "containsAll"
                )
            {
                // `first()` / `last()` THROW NoSuchElementException on an
                // empty collection; only the `OrNull` forms answer null. A
                // bare `[0]` read answered undefined for both.
                if matches!(field.as_str(), "first" | "firstOrNull") && args.is_empty() {
                    let target = if field == "first" {
                        "__kt_first"
                    } else {
                        "__kt_first_or_null"
                    };
                    *expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident(target)),
                        args: vec![Argument::positional(Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident("__kt_to_list")),
                            args: vec![Argument::positional((**object).clone())],
                            optional: false }))],
                        optional: false });
                    return;
                }
                if matches!(field.as_str(), "last" | "lastOrNull") && args.is_empty() {
                    let target = if field == "last" {
                        "__kt_last"
                    } else {
                        "__kt_last_or_null"
                    };
                    *expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident(target)),
                        args: vec![Argument::positional(Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident("__kt_to_list")),
                            args: vec![Argument::positional((**object).clone())],
                            optional: false }))],
                        optional: false });
                    return;
                }
                if field == "elementAt" && args.len() == 1 {
                    // `elementAt` THROWS IndexOutOfBoundsException out of
                    // range — a bare index read answered undefined.
                    *expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__kt_get_throwing")),
                        args: vec![
                            Argument::positional(Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::ident("__kt_to_list")),
                                args: vec![Argument::positional((**object).clone())],
                                optional: false })),
                            Argument::positional(args[0].value.clone()),
                        ],
                        optional: false });
                    return;
                }
                if field == "count" && args.len() == 1 {
                    *expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__coll_length")),
                        args: vec![Argument::positional(Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Member {
                                object: Box::new(Expression::new(ExprKind::Call {
                                    callee: Box::new(Expression::ident("__kt_to_list")),
                                    args: vec![Argument::positional((**object).clone())],
                                    optional: false })),
                                field: "filter".to_string(),
                                null_safe: false })),
                            args: args.clone(),
                            optional: false }))],
                        optional: false });
                    return;
                }
                if field == "average" && args.is_empty() {
                    let list = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__kt_to_list")),
                        args: vec![Argument::positional((**object).clone())],
                        optional: false });
                    let len = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__coll_length")),
                        args: vec![Argument::positional(list.clone())],
                        optional: false });
                    *expr = Expression::new(ExprKind::Ternary {
                        cond: Box::new(Expression::new(ExprKind::Binary {
                            op: BinOp::Eq,
                            left: Box::new(len.clone()),
                            right: Box::new(Expression::int(0)) })),
                        then: Box::new(Expression::new(ExprKind::Lit(Literal::Float(f64::NAN)))),
                        else_: Box::new(Expression::new(ExprKind::Binary {
                            op: BinOp::Div,
                            left: Box::new(Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::ident("__coll_sum")),
                                args: vec![Argument::positional(list)],
                                optional: false })),
                            right: Box::new(len) })) });
                    return;
                }
                if field == "toSortedSet" && args.is_empty() {
                    *expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__kt_to_set")),
                        args: vec![Argument::positional(Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident("__coll_sorted")),
                            args: vec![Argument::positional(Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::ident("__kt_to_list")),
                                args: vec![Argument::positional((**object).clone())],
                                optional: false }))],
                            optional: false }))],
                        optional: false });
                    return;
                }
                if field == "iterator" && args.is_empty() {
                    *expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__jvm_list_iterator")),
                        args: vec![Argument::positional(Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident("__kt_to_list")),
                            args: vec![Argument::positional((**object).clone())],
                            optional: false }))],
                        optional: false });
                    return;
                }
                if field == "none" && args.len() == 1 {
                    *expr = Expression::new(ExprKind::Unary {
                        op: UnaryOp::Not,
                        expr: Box::new(Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Member {
                                object: Box::new(Expression::new(ExprKind::Call {
                                    callee: Box::new(Expression::ident("__kt_to_list")),
                                    args: vec![Argument::positional((**object).clone())],
                                    optional: false })),
                                field: "any".to_string(),
                                null_safe: false })),
                            args: args.clone(),
                            optional: false })) });
                    return;
                }
                if field == "containsAll" && args.len() == 1 {
                    *expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__kt_contains_all")),
                        args: vec![Argument::positional((**object).clone()), args[0].clone()],
                        optional: false });
                    return;
                }
                *expr = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Member {
                        object: Box::new(Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident("__kt_to_list")),
                            args: vec![Argument::positional((**object).clone())],
                            optional: false })),
                        field: field.clone(),
                        null_safe: false })),
                    args: args.clone(),
                    optional: false });
                return;
            }
            // `zip(generateSequence(...))` — the sequence arrives inline, so
            // the VarDecl-time materialization never saw it; zip truncates by
            // the shorter side, so a bounded prefix serves the infinite case.
            if let ExprKind::Member { field, .. } = &callee.kind
                && field == "zip"
                && !args.is_empty()
            {
                if let Some(materialized) =
                    kotlin_materialize_generate_sequence(&args[0].value, Some(256))
                {
                    args[0].value = materialized;
                }
            }
            // ── Map receivers ───────────────────────────────────────────
            // Kotlin's Map overloads of the shared HOF spellings iterate
            // ENTRIES, and the `[array_methods]` HOF path intercepts these
            // member names before any adapter can. So a map-typed receiver
            // rewrites here: `filter`/`map` to the entry-aware adapters
            // (their results are a Map and a List respectively), the
            // predicate family to the same HOF over an ENTRY LIST view.
            if let ExprKind::Member { object, field, .. } = &callee.kind
                && kotlin_expr_type(object, locals, operators)
                    .as_deref()
                    .is_some_and(kotlin_type_is_plain_map_like)
            {
                match field.as_str() {
                    "filter" | "filterNot" if args.len() == 1 => {
                        let target = if field == "filter" {
                            "__kt_map_filter"
                        } else {
                            "__kt_map_filter_not"
                        };
                        *expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident(target)),
                            args: vec![Argument::positional((**object).clone()), args[0].clone()],
                            optional: false });
                        return;
                    }
                    "map" if args.len() == 1 => {
                        *expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident("__kt_map_to_list")),
                            args: vec![Argument::positional((**object).clone()), args[0].clone()],
                            optional: false });
                        return;
                    }
                    "any" | "all" | "none" | "count" | "forEach" | "find" | "first"
                    | "firstOrNull" | "minByOrNull" | "maxByOrNull" | "sumOf" | "flatMap"
                        if args.len() <= 1 =>
                    {
                        *expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Member {
                                object: Box::new(Expression::new(ExprKind::Call {
                                    callee: Box::new(Expression::ident("__kt_map_entry_list")),
                                    args: vec![Argument::positional((**object).clone())],
                                    optional: false })),
                                field: field.clone(),
                                null_safe: false })),
                            args: args.clone(),
                            optional: false });
                        return;
                    }
                    _ => {}
                }
            }
            // A `groupingBy { }` chain: its terminals `fold`/`reduce` would
            // otherwise be stolen by the array HOF table.
            if let ExprKind::Member { object, field, .. } = &callee.kind
                && kotlin_expr_type(object, locals, operators).as_deref() == Some("Grouping")
            {
                match field.as_str() {
                    "fold" if args.len() == 2 => {
                        // walk_expr already reversed `fold(init) { f }` into the
                        // JS reduce order `(f, init)`; the grouping adapter
                        // takes `(grouping, init, f)`.
                        *expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident("__kt_grouping_fold")),
                            args: vec![
                                Argument::positional((**object).clone()),
                                args[1].clone(),
                                args[0].clone(),
                            ],
                            optional: false });
                        return;
                    }
                    "reduce" if args.len() == 1 => {
                        *expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident("__kt_grouping_reduce")),
                            args: vec![Argument::positional((**object).clone()), args[0].clone()],
                            optional: false });
                        return;
                    }
                    _ => {}
                }
            }
            if let ExprKind::Ident(name) = &callee.kind {
                if name == "println"
                    && args.len() == 1
                    && kotlin_expr_type(&args[0].value, locals, operators)
                        .as_deref()
                        .is_some_and(kotlin_type_is_double_like)
                {
                    *expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__kt_print_double")),
                        args: args.clone(),
                        optional: false });
                    return;
                }
                // TreeSet's `add` keeps the backing list SORTED — the generic
                // add rewrite pushed in insertion order. Same re-route shape
                // as PriorityQueue below.
                if name == "__kt_add"
                    && args.len() == 2
                    && kotlin_expr_type(&args[0].value, locals, operators)
                        .as_deref()
                        .is_some_and(|ty| ty == "TreeSet")
                {
                    *expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Member {
                            object: Box::new(args[0].value.clone()),
                            field: "add".to_string(),
                            null_safe: false })),
                        args: vec![args[1].clone()],
                        optional: false });
                    return;
                }
                if name == "__coll_push" && args.len() == 2 {
                    if kotlin_expr_type(&args[0].value, locals, operators)
                        .as_deref()
                        .is_some_and(|ty| ty == "PriorityQueue")
                    {
                        *expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Member {
                                object: Box::new(args[0].value.clone()),
                                field: "add".to_string(),
                                null_safe: false })),
                            args: vec![args[1].clone()],
                            optional: false });
                        return;
                    }
                    if kotlin_expr_type(&args[0].value, locals, operators)
                        .as_deref()
                        .is_some_and(kotlin_type_is_set_like)
                    {
                        let value = args[1].value.clone();
                        *expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident("__kt_set_add")),
                            args: vec![
                                args[0].clone(),
                                Argument::positional(kotlin_key_expr(value.clone())),
                                Argument::positional(value),
                            ],
                            optional: false });
                        return;
                    }
                }
                if name == "__kt_to_set" && args.len() == 1 {
                    return;
                }
                // `groupingBy { }.reduce { }` — walk_expr rewrote `reduce` to
                // `__kt_reduce` before the receiver's type was knowable; a
                // Grouping receiver re-routes to the grouping terminal here.
                if name == "__kt_reduce"
                    && args.len() == 2
                    && kotlin_expr_type(&args[0].value, locals, operators).as_deref()
                        == Some("Grouping")
                {
                    *expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__kt_grouping_reduce")),
                        args: args.clone(),
                        optional: false });
                    return;
                }
                // `groupingBy { }.reduce { }` — walk_expr rewrote `reduce` to
                // `__kt_reduce` before the receiver's type was knowable; a
                // Grouping receiver re-routes to the grouping terminal here.
                if name == "__kt_reduce"
                    && args.len() == 2
                    && kotlin_expr_type(&args[0].value, locals, operators).as_deref()
                        == Some("Grouping")
                {
                    *expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__kt_grouping_reduce")),
                        args: args.clone(),
                        optional: false });
                    return;
                }
                if name == "__coll_length" && args.len() == 1 {
                    if kotlin_expr_type(&args[0].value, locals, operators)
                        .as_deref()
                        .is_some_and(kotlin_type_is_set_like)
                    {
                        *expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident("__kt_set_size")),
                            args: args.clone(),
                            optional: false });
                        return;
                    }
                    if kotlin_expr_type(&args[0].value, locals, operators)
                        .as_deref()
                        .is_some_and(kotlin_type_is_jvm_map_like)
                    {
                        *expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident("__jvm_map_size")),
                            args: args.clone(),
                            optional: false });
                        return;
                    }
                    if let Some(kind) = kotlin_expr_type(&args[0].value, locals, operators)
                        .and_then(|ty| kotlin_delegated_collection_kind(&ty))
                    {
                        if matches!(
                            kind.rsplit('.').next().unwrap_or(&kind),
                            "List"
                                | "MutableList"
                                | "Set"
                                | "MutableSet"
                                | "Collection"
                                | "Iterable"
                        ) {
                            *expr = Expression::new(ExprKind::Member {
                                object: Box::new(args[0].value.clone()),
                                field: "size".to_string(),
                                null_safe: false });
                            return;
                        }
                    }
                }
                if matches!(
                    name.as_str(),
                    "__coll_sum"
                        | "__coll_min"
                        | "__coll_max"
                        | "__coll_join"
                        | "__coll_slice"
                        | "__array_map"
                        | "__array_filter"
                        | "__array_reduce"
                        | "__array_some"
                        | "__array_every"
                ) && !args.is_empty()
                    && kotlin_expr_type(&args[0].value, locals, operators)
                        .as_deref()
                        .is_some_and(kotlin_type_is_set_like)
                {
                    let mut next_args = args.clone();
                    next_args[0].value = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__kt_to_list")),
                        args: vec![Argument::positional(args[0].value.clone())],
                        optional: false });
                    *expr = Expression::new(ExprKind::Call {
                        callee: callee.clone(),
                        args: next_args,
                        optional: false });
                    return;
                }
                if name == "__coll_contains" && args.len() == 2 {
                    if kotlin_expr_type(&args[0].value, locals, operators)
                        .as_deref()
                        .is_some_and(kotlin_type_is_set_like)
                    {
                        *expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident("__dict_has")),
                            args: vec![
                                args[0].clone(),
                                Argument::positional(kotlin_key_expr(args[1].value.clone())),
                            ],
                            optional: false });
                        return;
                    }
                }
                if name == "__dict_keys" && args.len() == 1 {
                    if kotlin_expr_type(&args[0].value, locals, operators)
                        .as_deref()
                        .is_some_and(kotlin_type_is_jvm_map_like)
                    {
                        *expr = Expression::new(ExprKind::Member {
                            object: Box::new(args[0].value.clone()),
                            field: "keys".to_string(),
                            null_safe: false });
                        return;
                    }
                    if let Some(kind) = kotlin_expr_type(&args[0].value, locals, operators)
                        .and_then(|ty| kotlin_delegated_collection_kind(&ty))
                    {
                        if matches!(
                            kind.rsplit('.').next().unwrap_or(&kind),
                            "Map" | "MutableMap"
                        ) {
                            *expr = Expression::new(ExprKind::Member {
                                object: Box::new(args[0].value.clone()),
                                field: "keys".to_string(),
                                null_safe: false });
                            return;
                        }
                    }
                }
                if name == "__dict_values" && args.len() == 1 {
                    if kotlin_expr_type(&args[0].value, locals, operators)
                        .as_deref()
                        .is_some_and(kotlin_type_is_jvm_map_like)
                    {
                        *expr = Expression::new(ExprKind::Member {
                            object: Box::new(args[0].value.clone()),
                            field: "values".to_string(),
                            null_safe: false });
                        return;
                    }
                }
                if name == "__coll_join" && args.len() == 2 {
                    if kotlin_expr_type(&args[0].value, locals, operators)
                        .and_then(|ty| kotlin_delegated_collection_kind(&ty))
                        .is_some()
                    {
                        *expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Member {
                                object: Box::new(args[0].value.clone()),
                                field: "joinToString".to_string(),
                                null_safe: false })),
                            args: vec![args[1].clone()],
                            optional: false });
                        return;
                    }
                }
                if let Some(reordered) = kotlin_normalized_fn_args(name, args) {
                    *args = reordered;
                }
                // Typed-array LITERAL factories materialize here (not at
                // walk time) so `val a = byteArrayOf(…)` types its local
                // first — that type drives identity `==` and rendering.
                if matches!(
                    name.as_str(),
                    "intArrayOf" | "doubleArrayOf" | "booleanArrayOf" | "charArrayOf"
                        | "longArrayOf" | "byteArrayOf" | "shortArrayOf" | "floatArrayOf"
                        | "uintArrayOf" | "ubyteArrayOf" | "ushortArrayOf" | "ulongArrayOf"
                ) {
                    *expr = Expression::new(ExprKind::Array(
                        args.iter()
                            .map(|arg| ArrayElement {
                                key: None,
                                value: arg.value.clone(),
                                spread: false,
                                by_ref: false })
                            .collect(),
                    ));
                    return;
                }
                // `Array(n) { init }` / `IntArray(n) { init }` / `List(n) { }`
                // — size + per-index initializer; a bare numeric-array ctor
                // (`IntArray(3)`) zero-fills. These parsed as plain calls to
                // an undefined name.
                if matches!(
                    name.as_str(),
                    "Array" | "IntArray" | "LongArray" | "DoubleArray" | "FloatArray"
                        | "CharArray" | "BooleanArray" | "ByteArray" | "ShortArray"
                        | "UIntArray" | "UByteArray" | "UShortArray" | "ULongArray"
                        | "List" | "MutableList"
                ) {
                    let with_init = args.len() == 2
                        && matches!(args[1].value.kind, ExprKind::Lambda { .. });
                    let bare_numeric = args.len() == 1
                        && name != "Array"
                        && name != "List"
                        && name != "MutableList"
                        && !matches!(args[0].value.kind, ExprKind::Lambda { .. });
                    if with_init || bare_numeric {
                        *expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident("__kt_array_init")),
                            args: args.clone(),
                            optional: false });
                        return;
                    }
                }
            }
            if let ExprKind::Member { object, field, .. } = &callee.kind
                && let Some(receiver_ty) = kotlin_expr_type(object, locals, operators)
            {
                // `x.toString()` on a statically Double value renders
                // Kotlin-style (`4.0`, never `4`).
                if field == "toString"
                    && args.is_empty()
                    && kotlin_type_is_double_like(&receiver_ty)
                {
                    *expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__kt_double_str")),
                        args: vec![Argument::positional(*object.clone())],
                        optional: false });
                    return;
                }
                if receiver_ty == "Vector" && matches!(field.as_str(), "iterator" | "listIterator")
                {
                    let mut call_args = vec![Argument::positional(*object.clone())];
                    call_args.extend(args.clone());
                    *expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__jvm_list_iterator")),
                        args: call_args,
                        optional: false });
                    return;
                }
                if receiver_ty == "java.util.Iterator" && field == "next" && args.is_empty() {
                    let list = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__dict_get")),
                        args: vec![
                            Argument::positional(*object.clone()),
                            Argument::positional(Expression::new(ExprKind::Lit(Literal::Str(
                                "__list".to_string(),
                            )))),
                        ],
                        optional: false });
                    let index = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__dict_get")),
                        args: vec![
                            Argument::positional(*object.clone()),
                            Argument::positional(Expression::new(ExprKind::Lit(Literal::Str(
                                "__index".to_string(),
                            )))),
                        ],
                        optional: false });
                    *expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__coll_get")),
                        args: vec![Argument::positional(list), Argument::positional(index)],
                        optional: false });
                    return;
                }
            }
        }
        ExprKind::New { class, args } => {
            normalize_kotlin_operator_expr(class, operators, locals);
            for arg in &mut *args {
                normalize_kotlin_operator_expr(&mut arg.value, operators, locals);
            }
            // Typed-array constructors also arrive as `New` (`UIntArray(4)
            // { … }` — uppercase heuristic); same lowering as the Call form.
            if let ExprKind::Ident(name) = &class.kind
                && matches!(
                    name.as_str(),
                    "Array" | "IntArray" | "LongArray" | "DoubleArray" | "FloatArray"
                        | "CharArray" | "BooleanArray" | "ByteArray" | "ShortArray"
                        | "UIntArray" | "UByteArray" | "UShortArray" | "ULongArray"
                )
                && !is_user_class_name(name)
            {
                let with_init = args.len() == 2
                    && matches!(args[1].value.kind, ExprKind::Lambda { .. });
                let bare_numeric =
                    args.len() == 1 && !matches!(args[0].value.kind, ExprKind::Lambda { .. });
                if with_init || bare_numeric {
                    *expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__kt_array_init")),
                        args: args.clone(),
                        optional: false });
                    return;
                }
            }
            if let ExprKind::Member { field, .. } = &class.kind
                && let Some(qualified) = qualified_inner_class(field)
            {
                **class = Expression::ident(&qualified);
            }
            if let ExprKind::Ident(name) = &class.kind {
                if is_qualified_inner_class_path(name) {
                    *expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident(name)),
                        args: args.clone(),
                        optional: false });
                    return;
                }
                if let Some(call) = kotlin_user_constructor_call(name, args) {
                    *expr = call;
                    return;
                }
                *args = kotlin_normalized_constructor_args(name, args);
            }
        }
        ExprKind::Member {
            object,
            field,
            null_safe } => {
            normalize_kotlin_operator_expr(object, operators, locals);
            if !*null_safe
                && field == "next"
                && let Some(index) =
                    kotlin_data_class_property_index(object, field, locals, operators)
            {
                *expr = Expression::new(ExprKind::Index {
                    object: object.clone(),
                    index: Box::new(Expression::int(index as i64)),
                    null_safe: false });
                return;
            }
        }
        ExprKind::Index { object, index, .. } => {
            normalize_kotlin_operator_expr(object, operators, locals);
            normalize_kotlin_operator_expr(index, operators, locals);
            if kotlin_expr_type(object, locals, operators)
                .as_deref()
                .is_some_and(kotlin_type_is_set_like)
            {
                *object = Box::new(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__kt_to_list")),
                    args: vec![Argument::positional((**object).clone())],
                    optional: false }));
                return;
            }
            if let Some(kind) = kotlin_expr_type(object, locals, operators)
                .and_then(|ty| kotlin_delegated_collection_kind(&ty))
            {
                if matches!(
                    kind.rsplit('.').next().unwrap_or(&kind),
                    "List" | "MutableList" | "Map" | "MutableMap"
                ) {
                    *expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Member {
                            object: object.clone(),
                            field: "get".to_string(),
                            null_safe: false })),
                        args: vec![Argument::positional((**index).clone())],
                        optional: false });
                    return;
                }
            }
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
                    lambda_locals.insert(param.name.clone(), ty.clone().to_string());
                }
            }
            match body {
                LambdaBody::Expr(expr) => {
                    normalize_kotlin_operator_expr(expr, operators, &lambda_locals)
                }
                LambdaBody::Block(stmts) => {
                    normalize_kotlin_operator_stmts(stmts, operators, &mut lambda_locals);
                }
            }
        }
        _ => {}
    }

    if let Some(replacement) = kotlin_static_array_map_rewrite(expr) {
        *expr = replacement;
        return;
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
            // `==` on the TYPED arrays is reference IDENTITY (List keeps
            // structural equals) — route around the structural value-eq slot.
            if matches!(op, BinOp::Eq | BinOp::NotEq)
                && kotlin_expr_type(left, locals, operators)
                    .as_deref()
                    .is_some_and(kotlin_type_is_typed_array)
                && kotlin_expr_type(right, locals, operators)
                    .as_deref()
                    .is_some_and(kotlin_type_is_typed_array)
            {
                let eq = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__kt_ref_eq")),
                    args: vec![
                        Argument::positional((**left).clone()),
                        Argument::positional((**right).clone()),
                    ],
                    optional: false });
                return Some(if *op == BinOp::Eq {
                    eq
                } else {
                    Expression::new(ExprKind::Unary {
                        op: UnaryOp::Not,
                        expr: Box::new(eq) })
                });
            }
            // Kotlin `==` is `.equals`: java.time values are `{sec, nano,
            // offset}` OBJECTS, and the generic eq slot answered reference
            // identity (two equal parses compared false). `compareTo` is the
            // type's own equality and already tree-dispatched.
            if matches!(op, BinOp::Eq | BinOp::NotEq)
                && kotlin_expr_type(left, locals, operators)
                    .as_deref()
                    .is_some_and(|t| t.starts_with("java.time."))
                && kotlin_expr_type(right, locals, operators)
                    .as_deref()
                    .is_some_and(|t| t.starts_with("java.time."))
            {
                let cmp = Expression::new(ExprKind::Binary {
                    op: if *op == BinOp::Eq { BinOp::Eq } else { BinOp::NotEq },
                    left: Box::new(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Member {
                            object: Box::new((**left).clone()),
                            field: "compareTo".to_string(),
                            null_safe: false })),
                        args: vec![Argument::positional((**right).clone())],
                        optional: false })),
                    right: Box::new(Expression::int(0)) });
                return Some(cmp);
            }
            if *op == BinOp::In {
                // A char range folds to its char array so membership works.
                let right_folded = {
                    let mut r = (**right).clone();
                    kotlin_fold_char_range(&mut r);
                    r
                };
                // `'t' in text` — Kotlin's `in` on a String is `contains`.
                if kotlin_expr_type(right, locals, operators)
                    .as_deref()
                    .map(|t| {
                        t.split('<').next().unwrap_or(t).rsplit('.').next().unwrap_or(t)
                            == "String"
                    })
                    .unwrap_or(false)
                {
                    return Some(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__kt_contains_ic")),
                        args: vec![
                            Argument::positional((**right).clone()),
                            Argument::positional((**left).clone()),
                            Argument::positional(Expression::new(ExprKind::Lit(
                                Literal::Bool(false),
                            ))),
                        ],
                        optional: false }));
                }
                if matches!(right.kind, ExprKind::Range { .. })
                    || matches!(
                        (&right.kind, &right_folded.kind),
                        (ExprKind::Range { .. }, _) | (_, ExprKind::Array(_))
                    ) && matches!(right.kind, ExprKind::Range { .. })
                    || kotlin_expr_type(right, locals, operators).as_deref() == Some("Range")
                {
                    return Some(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__coll_contains")),
                        args: vec![
                            Argument::positional(right_folded.clone()),
                            Argument::positional((**left).clone()),
                        ],
                        optional: false }));
                }
                if kotlin_expr_type(right, locals, operators)
                    .as_deref()
                    .is_some_and(kotlin_type_is_map_like)
                {
                    if kotlin_expr_type(right, locals, operators)
                        .as_deref()
                        .is_some_and(kotlin_type_is_set_like)
                    {
                        return Some(kotlin_set_contains_expr(
                            (**right).clone(),
                            (**left).clone(),
                        ));
                    }
                    return Some(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__dict_has")),
                        args: vec![
                            Argument::positional((**right).clone()),
                            Argument::positional(kotlin_key_expr((**left).clone())),
                        ],
                        optional: false }));
                }
                if let Some(kind) = kotlin_expr_type(right, locals, operators)
                    .and_then(|ty| kotlin_delegated_collection_kind(&ty))
                {
                    if matches!(
                        kind.rsplit('.').next().unwrap_or(&kind),
                        "List" | "MutableList" | "Set" | "MutableSet" | "Collection"
                    ) {
                        return Some(Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Member {
                                object: Box::new((**right).clone()),
                                field: "contains".to_string(),
                                null_safe: false })),
                            args: vec![Argument::positional((**left).clone())],
                            optional: false }));
                    }
                    if matches!(
                        kind.rsplit('.').next().unwrap_or(&kind),
                        "Map" | "MutableMap"
                    ) {
                        return Some(Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Member {
                                object: Box::new((**right).clone()),
                                field: "containsKey".to_string(),
                                null_safe: false })),
                            args: vec![Argument::positional((**left).clone())],
                            optional: false }));
                    }
                }
                let ty = kotlin_expr_type(right, locals, operators)?;
                let method = crate::protocol::binary_operator_method(*op)?;
                if operators.get(&ty).is_some_and(|info| info.has(method)) {
                    return Some(kotlin_operator_call(
                        (**right).clone(),
                        method,
                        vec![(**left).clone()],
                    ));
                }
                return None;
            }

            if kotlin_expr_type(left, locals, operators)
                .as_deref()
                .is_some_and(kotlin_type_is_set_like)
            {
                return match op {
                    BinOp::Add => Some(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__kt_set_union")),
                        args: vec![
                            Argument::positional((**left).clone()),
                            Argument::positional((**right).clone()),
                        ],
                        optional: false })),
                    BinOp::Sub => Some(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__kt_set_subtract")),
                        args: vec![
                            Argument::positional((**left).clone()),
                            Argument::positional((**right).clone()),
                        ],
                        optional: false })),
                    _ => None };
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
                _ => Some(call) }
        }
        ExprKind::Assign { target, value } => {
            // `holder.doubled = 10` where `doubled` is an extension property
            // WITH a setter — the write IS the setter call. The walk-time
            // read rewrite already turned the TARGET into `doubled(holder)`,
            // so the getter-call shape is matched too.
            let ext_set_target: Option<(String, Expression)> = match &target.kind {
                ExprKind::Member { object, field, .. }
                    if EXTENSION_PROPERTY_SETTERS
                        .with(|set| set.borrow().contains(field.as_str()))
                        && !is_user_property_name(field) =>
                {
                    Some((field.clone(), (**object).clone()))
                }
                ExprKind::Call { callee, args, .. }
                    if args.len() == 1
                        && matches!(&callee.kind, ExprKind::Ident(n)
                            if EXTENSION_PROPERTY_SETTERS
                                .with(|set| set.borrow().contains(n.as_str()))) =>
                {
                    let ExprKind::Ident(n) = &callee.kind else { unreachable!() };
                    Some((n.clone(), args[0].value.clone()))
                }
                _ => None };
            if let Some((prop, recv)) = ext_set_target {
                return Some(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident(&format!("{prop}__ext_set"))),
                    args: vec![
                        Argument::positional(recv),
                        Argument::positional((**value).clone()),
                    ],
                    optional: false }));
            }
            if let ExprKind::Index { object, index, .. } = &target.kind {
                if kotlin_expr_type(object, locals, operators)
                    .as_deref()
                    .is_some_and(kotlin_type_is_map_like)
                {
                    return Some(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__dict_set")),
                        args: vec![
                            Argument::positional((**object).clone()),
                            Argument::positional(kotlin_key_expr((**index).clone())),
                            Argument::positional((**value).clone()),
                        ],
                        optional: false }));
                }
                // `sb[i] = c` — StringBuilder index writes are `setCharAt`.
                if kotlin_expr_type(object, locals, operators).as_deref()
                    == Some("StringBuilder")
                {
                    return Some(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Member {
                            object: object.clone(),
                            field: "setCharAt".to_string(),
                            null_safe: false })),
                        args: vec![
                            Argument::positional((**index).clone()),
                            Argument::positional((**value).clone()),
                        ],
                        optional: false }));
                }
                // LIST index writes route through `__coll_set` so the
                // `[builtin_slots.array] set_item` binding applies —
                // `Collections.unmodifiableList` views must throw, and the
                // plain Index-assign path never consults the slot table for
                // arrays.
                if kotlin_expr_type(object, locals, operators)
                    .as_deref()
                    .is_some_and(kotlin_type_is_list_like)
                    && !matches!(index.kind, ExprKind::Range { .. } | ExprKind::Slice { .. })
                {
                    return Some(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__coll_set")),
                        args: vec![
                            Argument::positional((**object).clone()),
                            Argument::positional((**index).clone()),
                            Argument::positional((**value).clone()),
                        ],
                        optional: false }));
                }
            }
            let ExprKind::Binary { op, left, right } = &value.kind else {
                return None;
            };
            if !kotlin_same_simple_expr(target, left) {
                return None;
            }
            if kotlin_expr_type(target, locals, operators)
                .as_deref()
                .is_some_and(kotlin_type_is_set_like)
            {
                return match op {
                    BinOp::Add => Some(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__kt_add_all")),
                        args: vec![
                            Argument::positional((**target).clone()),
                            Argument::positional((**right).clone()),
                        ],
                        optional: false })),
                    BinOp::Sub => Some(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__kt_remove_all")),
                        args: vec![
                            Argument::positional((**target).clone()),
                            Argument::positional((**right).clone()),
                        ],
                        optional: false })),
                    _ => None };
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
            info.has(method).then(|| {
                Expression::new(ExprKind::Assign {
                    target: Box::new((**target).clone()),
                    value: Box::new(kotlin_operator_call((**target).clone(), method, Vec::new())) })
            })
        }
        ExprKind::Unary { op, expr: inner } => {
            let method = crate::protocol::unary_operator_method(*op)?;
            let ty = kotlin_expr_type(inner, locals, operators)?;
            operators
                .get(&ty)
                .is_some_and(|info| info.has(method))
                .then(|| kotlin_operator_call((**inner).clone(), method, Vec::new()))
        }
        ExprKind::Call { callee, args, .. } => {
            if args.len() == 1
                && let ExprKind::Member { object, field, .. } = &callee.kind
                && kotlin_expr_type(object, locals, operators)
                    .as_deref()
                    .is_some_and(kotlin_type_is_set_like)
            {
                let arg_is_lambda = matches!(args[0].value.kind, ExprKind::Lambda { .. });
                let helper = match field.as_str() {
                    "union" => Some("__kt_set_union"),
                    "intersect" => Some("__kt_set_intersect"),
                    "subtract" => Some("__kt_set_subtract"),
                    // `removeAll { p }` / `retainAll { p }` take a PREDICATE —
                    // the collection-argument adapters read the lambda as an
                    // element list and removed nothing.
                    "removeAll" if arg_is_lambda => Some("__kt_remove_if"),
                    "retainAll" if arg_is_lambda => Some("__kt_retain_if"),
                    "removeIf" if arg_is_lambda => Some("__kt_remove_if"),
                    "removeAll" => Some("__kt_remove_all"),
                    "retainAll" => Some("__kt_retain_all"),
                    _ => None };
                if let Some(helper) = helper {
                    return Some(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident(helper)),
                        args: vec![Argument::positional((**object).clone()), args[0].clone()],
                        optional: false }));
                }
            }
            // `Factory.from(7)` — an extension fn on the COMPANION: a free
            // call taking the class object as its receiver.
            if let ExprKind::Member { object, field, .. } = &callee.kind
                && let ExprKind::Ident(class) = &object.kind
                && COMPANION_EXT_FNS
                    .with(|set| set.borrow().contains(&(class.clone(), field.clone())))
            {
                let mut new_args = vec![Argument::positional((**object).clone())];
                new_args.extend(args.iter().cloned());
                return Some(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident(field)),
                    args: new_args,
                    optional: false }));
            }
            // `"a".upper()` where `upper` is a LOCAL holding a
            // receiver-typed function — the receiver is the first argument.
            if let ExprKind::Member { object, field, .. } = &callee.kind
                && locals
                    .get(field.as_str())
                    .is_some_and(|h| kotlin_hint_is_receiver_fn(h))
            {
                let mut new_args = vec![Argument::positional((**object).clone())];
                new_args.extend(args.iter().cloned());
                return Some(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident(field)),
                    args: new_args,
                    optional: false }));
            }
            // `runCatching { body }` — the body inlined under try/catch.
            if let ExprKind::Ident(name) = &callee.kind
                && name == "runCatching"
                && args.len() == 1
                && matches!(args[0].value.kind, ExprKind::Lambda { .. })
            {
                let ExprKind::Lambda { body, .. } = &args[0].value.kind else {
                    unreachable!()
                };
                let mut try_body: Vec<Statement> = Vec::new();
                match body {
                    LambdaBody::Expr(e) => try_body.push(Statement::new(StmtKind::Return(
                        Some(kotlin_result_obj(true, (**e).clone())),
                    ))),
                    LambdaBody::Block(stmts) => {
                        let mut stmts = stmts.clone();
                        match stmts.pop() {
                            // The walker may have wrapped the tail in a
                            // Return already — both spellings ARE the value.
                            Some(Statement { kind: StmtKind::Expr(e), .. })
                            | Some(Statement { kind: StmtKind::Return(Some(e)), .. }) => {
                                try_body.extend(stmts);
                                try_body.push(Statement::new(StmtKind::Return(Some(
                                    kotlin_result_obj(true, e),
                                ))));
                            }
                            Some(last) => {
                                try_body.extend(stmts);
                                try_body.push(last);
                                try_body.push(Statement::new(StmtKind::Return(Some(
                                    kotlin_result_obj(true, Expression::null()),
                                ))));
                            }
                            None => try_body.push(Statement::new(StmtKind::Return(Some(
                                kotlin_result_obj(true, Expression::null()),
                            )))) }
                    }
                }
                let catch = CatchClause {
                    types: Vec::new(),
                    var_name: Some("__kre".to_string()),
                    stack_var: None,
                    body: vec![Statement::new(StmtKind::Return(Some(kotlin_result_obj(
                        false,
                        Expression::ident("__kre"),
                    ))))],
                    when_clause: None };
                let wrapper = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Lambda {
                        params: Vec::new(),
                        body: LambdaBody::Block(vec![Statement::new(StmtKind::Try {
                            body: try_body,
                            catches: vec![catch],
                            else_body: None,
                            finally: None })]),
                        is_async: false,
                        captures: Vec::new() })),
                    args: Vec::new(),
                    optional: false });
                return Some(kotlin_as_result(wrapper));
            }
            // `Result.success(x)` / `Result.failure(e)`.
            if let ExprKind::Member { object, field, .. } = &callee.kind
                && matches!(&object.kind, ExprKind::Ident(n) if n == "Result")
                && matches!(field.as_str(), "success" | "failure")
                && args.len() == 1
            {
                return Some(kotlin_as_result(kotlin_result_obj(
                    field == "success",
                    args[0].value.clone(),
                )));
            }
            // `.fold`/`.map` were captured by `[array_methods]` at WALK time
            // (`__array_reduce`/`__array_map` free calls) — on a Result
            // receiver they are the Result combinators, put them back.
            if let ExprKind::Ident(name) = &callee.kind
                && let Some(member) = match name.as_str() {
                    "__array_reduce" if args.len() == 3 => Some("fold"),
                    "__array_map" | "__kt_map_hof" if args.len() == 2 => Some("map"),
                    _ => None }
                && kotlin_expr_type(&args[0].value, locals, operators)
                    .as_deref()
                    .is_some_and(kotlin_type_is_result)
            {
                return kotlin_operator_rewrite(
                    &Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Member {
                            object: Box::new(args[0].value.clone()),
                            field: member.to_string(),
                            null_safe: false })),
                        args: args[1..].to_vec(),
                        optional: false }),
                    operators,
                    locals,
                );
            }
            // Result combinators on a Result-typed receiver.
            if let ExprKind::Member { object, field, .. } = &callee.kind
                && kotlin_expr_type(object, locals, operators)
                    .as_deref()
                    .is_some_and(kotlin_type_is_result)
            {
                let recv = (**object).clone();
                let kr = || Expression::ident("__kr");
                let v = || kotlin_result_prop(Expression::ident("__kr"), "__kt_v");
                let e = || kotlin_result_prop(Expression::ident("__kr"), "__kt_e");
                let ret = |x: Expression| Statement::new(StmtKind::Return(Some(x)));
                match field.as_str() {
                    "getOrNull" if args.is_empty() => {
                        return Some(kotlin_result_iife(
                            recv,
                            kotlin_result_branch(true, vec![ret(v())], vec![ret(
                                Expression::null(),
                            )]),
                        ));
                    }
                    "exceptionOrNull" if args.is_empty() => {
                        return Some(kotlin_result_iife(
                            recv,
                            kotlin_result_branch(false, vec![ret(e())], vec![ret(
                                Expression::null(),
                            )]),
                        ));
                    }
                    "getOrThrow" if args.is_empty() => {
                        return Some(kotlin_result_iife(
                            recv,
                            kotlin_result_branch(true, vec![ret(v())], vec![
                                Statement::new(StmtKind::Throw {
                                    expr: Some(e()),
                                    cause: None }),
                            ]),
                        ));
                    }
                    "getOrDefault" if args.len() == 1 => {
                        return Some(kotlin_result_iife(
                            recv,
                            kotlin_result_branch(true, vec![ret(v())], vec![ret(
                                args[0].value.clone(),
                            )]),
                        ));
                    }
                    "getOrElse" if args.len() == 1 => {
                        let body = kotlin_result_inline_lambda(&args[0].value, e(), &|x| x);
                        return Some(kotlin_result_iife(
                            recv,
                            kotlin_result_branch(true, vec![ret(v())], body),
                        ));
                    }
                    "fold" if args.len() == 2 => {
                        // walk_expr already REVERSED two-arg `fold` into the
                        // JS reduce order (and dropped argument names), so
                        // `(onSuccess, onFailure)` arrives as
                        // `(onFailure, onSuccess)`.
                        let (on_ok, on_err) = (&args[1].value, &args[0].value);
                        let ok_body = kotlin_result_inline_lambda(on_ok, v(), &|x| x);
                        let err_body = kotlin_result_inline_lambda(on_err, e(), &|x| x);
                        return Some(kotlin_result_iife(
                            recv,
                            kotlin_result_branch(true, ok_body, err_body),
                        ));
                    }
                    "map" if args.len() == 1 => {
                        let body = kotlin_result_inline_lambda(&args[0].value, v(), &|x| {
                            kotlin_result_obj(true, x)
                        });
                        return Some(kotlin_as_result(kotlin_result_iife(
                            recv,
                            kotlin_result_branch(true, body, vec![ret(kr())]),
                        )));
                    }
                    "recover" if args.len() == 1 => {
                        let body = kotlin_result_inline_lambda(&args[0].value, e(), &|x| {
                            kotlin_result_obj(true, x)
                        });
                        return Some(kotlin_as_result(kotlin_result_iife(
                            recv,
                            kotlin_result_branch(false, body, vec![ret(kr())]),
                        )));
                    }
                    "mapCatching" | "recoverCatching" if args.len() == 1 => {
                        let on_ok = field == "mapCatching";
                        let arg = if on_ok { v() } else { e() };
                        let inner =
                            kotlin_result_inline_lambda(&args[0].value, arg, &|x| {
                                kotlin_result_obj(true, x)
                            });
                        let catch = CatchClause {
                            types: Vec::new(),
                            var_name: Some("__kre".to_string()),
                            stack_var: None,
                            body: vec![ret(kotlin_result_obj(
                                false,
                                Expression::ident("__kre"),
                            ))],
                            when_clause: None };
                        let guarded = vec![Statement::new(StmtKind::Try {
                            body: inner,
                            catches: vec![catch],
                            else_body: None,
                            finally: None })];
                        return Some(kotlin_as_result(kotlin_result_iife(
                            recv,
                            kotlin_result_branch(on_ok, guarded, vec![ret(kr())]),
                        )));
                    }
                    "onSuccess" | "onFailure" if args.len() == 1 => {
                        let on_ok = field == "onSuccess";
                        let arg = if on_ok { v() } else { e() };
                        // Side-effect body: inline plainly, the Return after
                        // the If answers the receiver for chaining.
                        let mut body =
                            kotlin_result_inline_lambda(&args[0].value, arg, &|x| x);
                        // The inline helper ends with `Return(tail)` — the
                        // combinator DISCARDS the lambda value: demote it.
                        if let Some(Statement { kind: StmtKind::Return(Some(tail)), .. }) =
                            body.pop()
                        {
                            body.push(Statement::new(StmtKind::Expr(tail)));
                        }
                        return Some(kotlin_as_result(kotlin_result_iife(
                            recv,
                            kotlin_result_branch(on_ok, body, vec![ret(kr())]),
                        )));
                    }
                    _ => {}
                }
            }
            // kotlin.time.Duration — spellings over a plain millisecond
            // number.
            if let ExprKind::Member { object, field, .. } = &callee.kind {
                // `N.toDuration(DurationUnit.X)` → ms.
                if field == "toDuration"
                    && args.len() == 1
                    && let Some((num, den)) = kotlin_duration_unit_factor(&args[0].value)
                {
                    // ms = N * den / num (the factor maps ms → unit).
                    let mut e = (**object).clone();
                    if den != 1.0 {
                        e = Expression::new(ExprKind::Binary {
                            op: BinOp::Mul,
                            left: Box::new(e),
                            right: Box::new(Expression::new(ExprKind::Lit(Literal::Float(
                                den,
                            )))) });
                    }
                    if num != 1.0 {
                        e = Expression::new(ExprKind::Binary {
                            op: BinOp::Div,
                            left: Box::new(e),
                            right: Box::new(Expression::new(ExprKind::Lit(Literal::Float(
                                num,
                            )))) });
                    }
                    return Some(e);
                }
                // `d.toLong(unit)` / `d.toInt(unit)` truncate toward zero;
                // `d.toDouble(unit)` divides exactly.
                if matches!(field.as_str(), "toLong" | "toInt" | "toDouble")
                    && args.len() == 1
                    && let Some((num, den)) = kotlin_duration_unit_factor(&args[0].value)
                {
                    let obj = (**object).clone();
                    if field == "toDouble" {
                        let mut e = obj;
                        if num != 1.0 {
                            e = Expression::new(ExprKind::Binary {
                                op: BinOp::Mul,
                                left: Box::new(e),
                                right: Box::new(Expression::new(ExprKind::Lit(
                                    Literal::Float(num),
                                ))) });
                        }
                        if den != 1.0 {
                            e = Expression::new(ExprKind::Binary {
                                op: BinOp::Div,
                                left: Box::new(e),
                                right: Box::new(Expression::new(ExprKind::Lit(
                                    Literal::Float(den),
                                ))) });
                        }
                        // Identity marker so the result TYPES as Double
                        // (the bare division re-typed as Duration and the
                        // toString formatter reformatted it).
                        return Some(Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident("__kt_as_double")),
                            args: vec![Argument::positional(e)],
                            optional: false }));
                    }
                    return Some(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__kt_dur_whole")),
                        args: vec![
                            Argument::positional(obj),
                            Argument::positional(Expression::new(ExprKind::Lit(
                                Literal::Float(num),
                            ))),
                            Argument::positional(Expression::new(ExprKind::Lit(
                                Literal::Float(den),
                            ))),
                        ],
                        optional: false }));
                }
                // `isNegative()` on a Duration is a plain sign test.
                if field == "isNegative"
                    && args.is_empty()
                    && matches!(
                        kotlin_expr_type(object, locals, operators).as_deref(),
                        Some("Duration") | Some("java.time.Duration")
                    )
                {
                    return Some(Expression::new(ExprKind::Binary {
                        op: BinOp::Lt,
                        left: Box::new((**object).clone()),
                        right: Box::new(Expression::int(0)) }));
                }
                // `toString()` on a kotlin.time Duration formats components.
                if field == "toString"
                    && args.is_empty()
                    && kotlin_expr_type(object, locals, operators).as_deref()
                        == Some("Duration")
                {
                    return Some(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__kt_duration_str")),
                        args: vec![Argument::positional((**object).clone())],
                        optional: false }));
                }
            }
            // A Duration SUM's type only this walker knows (the compiler
            // types Binary as a number) — answer `isZero` right here.
            if args.is_empty()
                && let ExprKind::Member { object, field, .. } = &callee.kind
                && field == "isZero"
                && matches!(
                    kotlin_expr_type(object, locals, operators).as_deref(),
                    Some("java.time.Duration")
                )
                && matches!(object.kind, ExprKind::Binary { .. })
            {
                return Some(Expression::new(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new((**object).clone()),
                    right: Box::new(Expression::int(0)) }));
            }
            // `LocalTime.second` was swallowed at WALK time by the Pair
            // `.second` probe (receiver types are unknown there). Now the
            // receiver IS typed: a java.* receiver's field goes back to a
            // real member read so the tree dispatches it.
            if let ExprKind::Ident(name) = &callee.kind
                && name == "__kt_tuple_prop"
                && args.len() == 3
                && kotlin_expr_type(&args[0].value, locals, operators)
                    .as_deref()
                    .is_some_and(|t| t.starts_with("java."))
                && let ExprKind::Lit(Literal::Str(field)) = &args[1].value.kind
            {
                return Some(Expression::new(ExprKind::Member {
                    object: Box::new(args[0].value.clone()),
                    field: field.to_string(),
                    null_safe: false }));
            }
            None
        }
        // `(x?.method { … }) ?: fallback` — the safe call lowered to
        // `if (x == null) null else call`; folding the coalesce INTO the
        // guard keeps a Unit-returning lambda from tripping the fallback
        // (our Unit is undefined, which the coalesce reads as absent).
        ExprKind::NullCoalesce { left, right } => {
            // ONLY when the guarded call invokes a LAMBDA (the scope-function
            // shapes) — `xs?.firstOrNull() ?: d` must still fall back when
            // the CALL itself answers null.
            if let ExprKind::Ternary { cond, then, else_ } = &left.kind
                && matches!(then.kind, ExprKind::Lit(Literal::Null))
                && matches!(
                    &else_.kind,
                    ExprKind::Call { callee, .. }
                        if matches!(callee.kind, ExprKind::Lambda { .. })
                )
            {
                return Some(Expression::new(ExprKind::Ternary {
                    cond: cond.clone(),
                    then: Box::new((**right).clone()),
                    else_: else_.clone() }));
            }
            None
        }
        ExprKind::Index { object, index, .. } => {
            let ty = kotlin_expr_type(object, locals, operators);
            // WALK time turned `.get(x)` into an Index before types existed;
            // on a java.time receiver `get(field)` is the temporal-field
            // accessor — put the member call back for tree dispatch.
            if ty.as_deref().is_some_and(|t| t.starts_with("java.time.")) {
                return Some(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Member {
                        object: Box::new((**object).clone()),
                        field: "get".to_string(),
                        null_safe: false })),
                    args: vec![Argument::positional((**index).clone())],
                    optional: false }));
            }
            if ty.as_deref().is_some_and(kotlin_type_is_map_like) {
                return Some(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__dict_get")),
                    args: vec![
                        Argument::positional((**object).clone()),
                        Argument::positional(kotlin_key_expr((**index).clone())),
                    ],
                    optional: false }));
            }
            // A LIST index read routes through `__coll_get` so the
            // `[builtin_slots.array] get_item` binding applies — Kotlin
            // throws IndexOutOfBoundsException out of range, a bare Index
            // answers undefined. Ranges/slices keep the plain Index path.
            if ty.as_deref().is_some_and(kotlin_type_is_list_like)
                && !matches!(
                    index.kind,
                    ExprKind::Range { .. } | ExprKind::Slice { .. }
                )
            {
                return Some(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__coll_get")),
                    args: vec![
                        Argument::positional((**object).clone()),
                        Argument::positional((**index).clone()),
                    ],
                    optional: false }));
            }
            None
        }
        _ => None }
}

/// List spellings for the `+`/`-` operator gate — `listOf(1) + listOf(2)`
/// concatenates, it never becomes numeric addition.
fn kotlin_type_is_list_like(ty: &str) -> bool {
    let bare = ty
        .split('<')
        .next()
        .unwrap_or(ty)
        .rsplit('.')
        .next()
        .unwrap_or(ty)
        .trim();
    matches!(
        bare,
        "List" | "MutableList" | "ArrayList" | "Collection" | "Iterable" | "Array"
            | "IntArray" | "LongArray" | "DoubleArray" | "FloatArray" | "BooleanArray"
            | "CharArray" | "ByteArray" | "ShortArray" | "UIntArray" | "UByteArray"
            | "UShortArray" | "ULongArray"
    )
}

/// The fixed-size ARRAY types — `==` on these is REFERENCE identity in
/// Kotlin, unlike List's structural equals.
fn kotlin_type_is_typed_array(ty: &str) -> bool {
    let bare = ty
        .split('<')
        .next()
        .unwrap_or(ty)
        .rsplit('.')
        .next()
        .unwrap_or(ty)
        .trim();
    matches!(
        bare,
        "IntArray" | "LongArray" | "DoubleArray" | "FloatArray" | "BooleanArray"
            | "CharArray" | "ByteArray" | "ShortArray" | "UIntArray" | "UByteArray"
            | "UShortArray" | "ULongArray"
    )
}

/// Map spellings only — `kotlin_type_is_map_like` also answers for Sets,
/// which have their own rewrite arm and must not take the entry-based one.
fn kotlin_type_is_plain_map_like(ty: &str) -> bool {
    kotlin_type_is_map_like(ty) && !kotlin_type_is_set_like(ty)
}

fn kotlin_type_is_map_like(ty: &str) -> bool {
    let bare = ty
        .split('<')
        .next()
        .unwrap_or(ty)
        .rsplit('.')
        .next()
        .unwrap_or(ty)
        .trim();
    matches!(
        bare,
        "Map"
            | "MutableMap"
            | "HashMap"
            | "LinkedHashMap"
            | "TreeMap"
            | "ConcurrentHashMap"
            | "IdentityHashMap"
            | "WeakHashMap"
            | "Hashtable"
            | "Properties"
            | "Set"
            | "MutableSet"
            | "HashSet"
            | "LinkedHashSet"
            | "TreeSet"
    )
}

fn kotlin_type_is_set_like(ty: &str) -> bool {
    let bare = ty
        .split('<')
        .next()
        .unwrap_or(ty)
        .rsplit('.')
        .next()
        .unwrap_or(ty)
        .trim();
    matches!(
        bare,
        "Set" | "MutableSet" | "HashSet" | "LinkedHashSet" | "TreeSet"
    )
}

fn kotlin_type_is_jvm_map_like(ty: &str) -> bool {
    let bare = ty
        .split('<')
        .next()
        .unwrap_or(ty)
        .rsplit('.')
        .next()
        .unwrap_or(ty)
        .trim();
    matches!(
        bare,
        "HashMap"
            | "LinkedHashMap"
            | "TreeMap"
            | "ConcurrentHashMap"
            | "IdentityHashMap"
            | "WeakHashMap"
            | "Hashtable"
            | "Properties"
    )
}

fn kotlin_delegated_collection_kind(ty: &str) -> Option<String> {
    let bare = ty.rsplit('.').next().unwrap_or(ty);
    KOTLIN_DELEGATED_COLLECTIONS.with(|map| map.borrow().get(bare).cloned())
}

fn kotlin_literal_keyed_collection_type(expr: &Expression) -> Option<&'static str> {
    let ExprKind::Object(props) = &expr.kind else {
        return None;
    };
    if props.iter().any(|prop| {
        matches!(
            prop,
            ObjectProperty::KeyValue { key, .. }
                if matches!(&key.kind, ExprKind::Lit(Literal::Str(s)) if s == SET_MARKER)
        )
    }) {
        Some("Set")
    } else {
        Some("Map")
    }
}

fn kotlin_generated_dict_items_source(expr: &Expression) -> Option<Expression> {
    let ExprKind::Ternary { cond, else_, .. } = &expr.kind else {
        return None;
    };
    let ExprKind::Call {
        callee: cond_callee,
        args: cond_args,
        ..
    } = &cond.kind
    else {
        return None;
    };
    if !matches!(&cond_callee.kind, ExprKind::Ident(name) if name == "__coll_is_array")
        || cond_args.len() != 1
    {
        return None;
    }
    let ExprKind::Call {
        callee: else_callee,
        args: else_args,
        ..
    } = &else_.kind
    else {
        return None;
    };
    if !matches!(&else_callee.kind, ExprKind::Ident(name) if name == "__dict_items")
        || else_args.len() != 1
    {
        return None;
    }
    Some(cond_args[0].value.clone())
}

fn kotlin_data_class_property_index(
    object: &Expression,
    field: &str,
    locals: &KotlinLocalTypes,
    operators: &KotlinOperatorTable,
) -> Option<usize> {
    let ty = kotlin_expr_type(object, locals, operators)?;
    KOTLIN_DATA_CLASS_PROPERTY_INDEX.with(|map| {
        map.borrow()
            .get(ty.rsplit('.').next().unwrap_or(&ty))
            .and_then(|fields| fields.get(field).copied())
    })
}

fn kotlin_key_expr(expr: Expression) -> Expression {
    match expr.kind {
        ExprKind::Lit(Literal::Str(_)) => expr,
        ExprKind::Lit(Literal::Int(value)) => {
            Expression::new(ExprKind::Lit(Literal::Str(value.to_string())))
        }
        ExprKind::Lit(Literal::Float(value)) => {
            Expression::new(ExprKind::Lit(Literal::Str(value.to_string())))
        }
        ExprKind::Lit(Literal::Bool(value)) => {
            Expression::new(ExprKind::Lit(Literal::Str(value.to_string())))
        }
        ExprKind::Lit(Literal::Null) => {
            Expression::new(ExprKind::Lit(Literal::Str("null".to_string())))
        }
        _ => Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__kt_tostring")),
            args: vec![Argument::positional(expr)],
            optional: false }) }
}

fn kotlin_dict_get_expr(object: Expression, key: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("__dict_get")),
        args: vec![Argument::positional(object), Argument::positional(key)],
        optional: false })
}

fn kotlin_dict_has_expr(object: Expression, key: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("__dict_has")),
        args: vec![Argument::positional(object), Argument::positional(key)],
        optional: false })
}

fn kotlin_set_contains_expr(object: Expression, value: Expression) -> Expression {
    let key = kotlin_key_expr(value.clone());
    Expression::new(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(kotlin_dict_has_expr(object.clone(), key.clone())),
        right: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(kotlin_dict_get_expr(object, key)),
            right: Box::new(value) })) })
}

fn kotlin_same_simple_expr(a: &Expression, b: &Expression) -> bool {
    match (&a.kind, &b.kind) {
        (ExprKind::Ident(a), ExprKind::Ident(b)) => a == b,
        (
            ExprKind::Member {
                object: ao,
                field: af,
                null_safe: ans },
            ExprKind::Member {
                object: bo,
                field: bf,
                null_safe: bns },
        ) => af == bf && ans == bns && kotlin_same_simple_expr(ao, bo),
        (
            ExprKind::Index {
                object: ao,
                index: ai,
                null_safe: ans },
            ExprKind::Index {
                object: bo,
                index: bi,
                null_safe: bns },
        ) => ans == bns && kotlin_same_simple_expr(ao, bo) && kotlin_same_simple_expr(ai, bi),
        (ExprKind::Lit(Literal::Int(a)), ExprKind::Lit(Literal::Int(b))) => a == b,
        (ExprKind::Lit(Literal::Str(a)), ExprKind::Lit(Literal::Str(b))) => a == b,
        _ => false }
}

fn kotlin_operator_call(receiver: Expression, method: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(receiver),
            field: method.to_string(),
            null_safe: false })),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false })
}

fn kotlin_compare_zero_call(helper: &str, value: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(helper)),
        args: vec![Argument::positional(value)],
        optional: false })
}

fn kotlin_type_is_double_like(ty: &str) -> bool {
    matches!(
        ty.trim()
            .trim_end_matches('?')
            .rsplit('.')
            .next()
            .unwrap_or(ty),
        "Double" | "Float" | "double" | "float"
    )
}

fn kotlin_expr_type(
    expr: &Expression,
    locals: &KotlinLocalTypes,
    operators: &KotlinOperatorTable,
) -> Option<String> {
    match &expr.kind {
        // A literal IS its type — `val s = "x"` must type `s` as String or
        // every literal-initialized local misses the string-aware rewrites
        // (and `var i = 1` must land in the locals map at all, or the
        // loop-variable shadow rename never sees it).
        ExprKind::Lit(Literal::Str(_)) => Some("String".to_string()),
        ExprKind::Lit(Literal::Int(_)) => Some("Int".to_string()),
        ExprKind::Lit(Literal::Float(_)) => Some("Double".to_string()),
        ExprKind::Lit(Literal::Bool(_)) => Some("Boolean".to_string()),
        ExprKind::Ident(name) => KOTLIN_KEYED_COLLECTION_TYPES
            .with(|map| map.borrow().get(name).cloned())
            .or_else(|| locals.get(name).cloned()),
        ExprKind::New { class, .. } => match &class.kind {
            ExprKind::Ident(name) => Some(name.clone()),
            ExprKind::Member { field, .. } => Some(field.clone()),
            _ => None },
        // Element reads keep the element type where it matters for
        // RENDERING: `println(floatArr[0])` must print `7.0`.
        ExprKind::Index { object, .. } => {
            match kotlin_expr_type(object, locals, operators).as_deref() {
                Some("FloatArray") | Some("DoubleArray") => Some("Double".to_string()),
                _ => None }
        }
        // Property reads on java.* receivers type through the tree too —
        // `value.dayOfWeek` is a DayOfWeek by jvm declaration.
        ExprKind::Member { object, field, .. } => {
            // `Duration.ZERO`/`Duration.INFINITE` — kotlin.time's own
            // constants (numbers at runtime, Duration to the rewrites).
            if matches!(&object.kind, ExprKind::Ident(n) if n == "Duration")
                && matches!(field.as_str(), "ZERO" | "INFINITE")
            {
                return Some("Duration".to_string());
            }
            let recv = kotlin_expr_type(object, locals, operators)?;
            if !recv.starts_with("java.") {
                return None;
            }
            vybe_runtime::namespaces::lookup_type_member_return(
                &["jvm".to_string(), "kotlin".to_string()],
                &recv,
                field,
            )
        }
        // The typed-array factories are erased to Array literals at parse
        // time; a FLOAT literal element is the surviving evidence that this
        // is a double/float collection (drives `sum()`/`[i]` rendering).
        ExprKind::Array(elements) => {
            if !elements.is_empty()
                && elements
                    .iter()
                    .any(|e| matches!(e.value.kind, ExprKind::Lit(Literal::Float(_))))
            {
                Some("DoubleArray".to_string())
            } else {
                None
            }
        }
        ExprKind::Call { callee, .. } => {
            if let ExprKind::Member { object, field, .. } = &callee.kind {
                if field == "groupingBy" {
                    return Some("Grouping".to_string());
                }
                // Double-returning members, so `println(arr.average())`
                // renders `2.0` (the Double-print rewrite needs a type).
                if matches!(field.as_str(), "average" | "toDouble" | "toFloat") {
                    return Some("Double".to_string());
                }
                if field == "toDuration" {
                    return Some("Duration".to_string());
                }
                if matches!(field.as_str(), "success" | "failure")
                    && matches!(&object.kind, ExprKind::Ident(n) if n == "Result")
                {
                    return Some("Result".to_string());
                }
                // `sum()`/`min()`/`max()` keep the element type — Double on
                // the float/double arrays.
                if matches!(
                    field.as_str(),
                    "sum" | "min" | "max" | "minOrNull" | "maxOrNull"
                ) && matches!(
                    kotlin_expr_type(object, locals, operators).as_deref(),
                    Some("FloatArray") | Some("DoubleArray")
                ) {
                    return Some("Double".to_string());
                }
                // A member call on a java.* receiver: the return type is
                // TREE DATA (`b.minusSeconds(1)` is a LocalTime because the
                // jvm platform says so), same resolver as dispatch.
                if let Some(recv) = kotlin_expr_type(object, locals, operators)
                    && recv.starts_with("java.")
                    && let Some(ret) = vybe_runtime::namespaces::lookup_type_member_return(
                        &["jvm".to_string(), "kotlin".to_string()],
                        &recv,
                        field,
                    )
                {
                    return Some(ret);
                }

            }
            if let ExprKind::Ident(name) = &callee.kind {
                // `sum`/`min`/`max` arrive here as `__coll_*` free calls (the
                // walk-time member rewrite) — keep the Double element type.
                if name == "__kt_as_double" {
                    return Some("Double".to_string());
                }
                if name == "runCatching" || name == "__kt_as_result" {
                    return Some("Result".to_string());
                }
                if matches!(
                    name.as_str(),
                    "__coll_sum" | "__coll_min" | "__coll_max" | "__coll_get"
                        | "__kt_get_throwing"
                ) && let ExprKind::Call { args, .. } = &expr.kind
                    && args.first().is_some_and(|a| {
                        matches!(
                            kotlin_expr_type(&a.value, locals, operators).as_deref(),
                            Some("FloatArray") | Some("DoubleArray")
                        )
                    })
                {
                    return Some("Double".to_string());
                }
                return match name.as_str() {
                    "mapOf" | "mutableMapOf" | "linkedMapOf" | "hashMapOf" | "buildMap"
                    | "emptyMap" => Some("Map".to_string()),
                    "setOf" | "mutableSetOf" | "linkedSetOf" | "hashSetOf" | "buildSet"
                    | "emptySet" | "__kt_to_set" | "__kt_set_union" | "__kt_set_intersect"
                    | "__kt_set_subtract" => Some("Set".to_string()),
                    "__kt_to_list" | "listOf" | "mutableListOf" | "arrayListOf"
                    | "emptyList" | "buildList" => Some("List".to_string()),
                    // The ctor spellings arrive as plain CALLS too.
                    "IntArray" | "LongArray" | "DoubleArray" | "FloatArray" | "BooleanArray"
                    | "CharArray" | "ByteArray" | "ShortArray" | "UIntArray" | "UByteArray"
                    | "UShortArray" | "ULongArray" => Some(name.clone()),
                    "intArrayOf" => Some("IntArray".to_string()),
                    "doubleArrayOf" => Some("DoubleArray".to_string()),
                    "booleanArrayOf" => Some("BooleanArray".to_string()),
                    "charArrayOf" => Some("CharArray".to_string()),
                    "longArrayOf" => Some("LongArray".to_string()),
                    "byteArrayOf" => Some("ByteArray".to_string()),
                    "shortArrayOf" => Some("ShortArray".to_string()),
                    "floatArrayOf" => Some("FloatArray".to_string()),
                    "uintArrayOf" => Some("UIntArray".to_string()),
                    "ubyteArrayOf" => Some("UByteArray".to_string()),
                    "ushortArrayOf" => Some("UShortArray".to_string()),
                    "ulongArrayOf" => Some("ULongArray".to_string()),
                    "__jvm_list_iterator" => Some("java.util.Iterator".to_string()),
                    _ => None };
            }
            if let Some(path) = dotted_expr_path(callee) {
                let lower = path.to_ascii_lowercase();
                if matches!(
                    lower.as_str(),
                    "java.util.collections.emptymap" | "java.util.collections.singletonmap"
                ) {
                    return Some("java.util.HashMap".to_string());
                }
                if lower == "java.util.collections.newsetfrommap" {
                    return Some("java.util.HashSet".to_string());
                }
                // The unmodifiable views keep the wrapped shape's type, so
                // the checked set_item/get_item slots fire on index writes.
                if matches!(
                    lower.as_str(),
                    "java.util.collections.unmodifiablelist"
                        | "java.util.collections.singletonlist"
                        | "java.util.collections.singleton"
                ) {
                    return Some("List".to_string());
                }
                if lower == "java.util.collections.unmodifiablemap" {
                    return Some("Map".to_string());
                }
                if lower == "java.util.collections.unmodifiableset" {
                    return Some("Set".to_string());
                }
                // Every java.* static's return type is TREE DATA —
                // `platforms/jvm` registers `member_returns` on its type
                // nodes, and this asks the common resolver instead of
                // keeping a walker-local copy of JVM facts. (The
                // `java.util.Collections` special cases above stay: they map
                // to KOTLIN shapes — `List`, `Map` — for this walker's own
                // rewrites, not to the java type the tree declares.)
                if let Some((class_path, member)) = path.rsplit_once('.')
                    && let Some(ret) = vybe_runtime::namespaces::lookup_type_member_return(
                        &["jvm".to_string(), "kotlin".to_string()],
                        class_path,
                        member,
                    )
                {
                    return Some(ret);
                }
                if matches!(
                    lower.as_str(),
                    "kotlin.math.sign"
                        | "kotlin.math.sqrt"
                        | "kotlin.math.pow"
                        | "kotlin.math.round"
                        | "kotlin.math.ceil"
                        | "kotlin.math.floor"
                        | "kotlin.math.sin"
                        | "kotlin.math.cos"
                        | "kotlin.math.tan"
                        | "kotlin.math.asin"
                        | "kotlin.math.acos"
                        | "kotlin.math.atan"
                        | "kotlin.math.atan2"
                        | "kotlin.math.sinh"
                        | "kotlin.math.cosh"
                        | "kotlin.math.tanh"
                        | "kotlin.math.exp"
                        | "kotlin.math.ln"
                        | "kotlin.math.log10"
                        | "kotlin.math.hypot"
                        | "kotlin.math.max"
                        | "kotlin.math.min"
                        | "kotlin.math.ulp"
                        | "kotlin.math.nextafter"
                        | "kotlin.math.nextup"
                        | "kotlin.math.nextdown"
                ) {
                    return Some("Double".to_string());
                }
            }
            if let ExprKind::Member { object, field, .. } = &callee.kind {
                let receiver_ty = kotlin_expr_type(object, locals, operators)?;
                if receiver_ty == "Vector" && matches!(field.as_str(), "iterator" | "listIterator")
                {
                    return Some("java.util.Iterator".to_string());
                }
                return operators
                    .get(&receiver_ty)
                    .and_then(|info| info.return_type(field));
            }
            None
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
            // `map - key` / `list + elem` keep the COLLECTION's type — the
            // right-operand fallback typed `mapOf(…) - "b"` as String, which
            // routed `.size` through the string paths.
            if matches!(op, BinOp::Add | BinOp::Sub)
                && (kotlin_type_is_map_like(&receiver_ty)
                    || kotlin_type_is_list_like(&receiver_ty))
            {
                return Some(receiver_ty);
            }
            // Duration arithmetic keeps the Duration — `(d + x).isZero()` /
            // `(base * 5).toLong(…)` need the result typed, or dispatch
            // falls to a bare number.
            if matches!(op, BinOp::Add | BinOp::Sub | BinOp::Mul | BinOp::Div)
                && (receiver_ty == "java.time.Duration" || receiver_ty == "Duration")
            {
                return Some(receiver_ty);
            }
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
            // `-(2.toDuration(…))` stays a Duration.
            if matches!(op, UnaryOp::Neg)
                && (receiver_ty == "Duration" || receiver_ty == "java.time.Duration")
            {
                return Some(receiver_ty);
            }
            operators
                .get(&receiver_ty)
                .and_then(|info| info.return_type(method))
        }
        ExprKind::Range { .. } => Some("Range".to_string()),
        ExprKind::Object(props) => {
            if props.iter().any(|prop| {
                matches!(
                    prop,
                    ObjectProperty::KeyValue { key, .. }
                        if matches!(&key.kind, ExprKind::Lit(Literal::Str(s)) if s == SET_MARKER)
                )
            }) {
                Some("Set".to_string())
            } else {
                Some("Map".to_string())
            }
        }
        _ => None }
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
                            spread: false });
                    }
                }
            }
            _ => {}
        }
    }

    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(&type_name)),
        args,
        optional: false })
}

/// True when `expr` is a bare dotted chain of identifiers (`java.util`), which
/// is what distinguishes a package-qualified type name from member access on a
/// value.
fn is_ident_chain(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Ident(_) => true,
        ExprKind::Member { object, .. } => is_ident_chain(object),
        _ => false }
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
            is_nullable: false }],
        body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Call {
            callee: Box::new(target),
            args: vec![Argument::positional(Expression::ident(&arg_name))],
            optional: false }))),
        is_async: false,
        captures: Vec::new() })
}

fn kotlin_class_literal_expr(name: &str) -> Expression {
    let simple = name.rsplit('.').next().unwrap_or(name);
    let java = Expression::new(ExprKind::Object(vec![
        ObjectProperty::KeyValue {
            key: Expression::string("name"),
            value: Expression::string(name) },
        ObjectProperty::KeyValue {
            key: Expression::string("canonicalName"),
            value: Expression::string(name) },
        ObjectProperty::KeyValue {
            key: Expression::string("simpleName"),
            value: Expression::string(simple) },
    ]));

    Expression::new(ExprKind::Object(vec![
        ObjectProperty::KeyValue {
            key: Expression::string("simpleName"),
            value: Expression::string(simple) },
        ObjectProperty::KeyValue {
            key: Expression::string("qualifiedName"),
            value: Expression::string(name) },
        ObjectProperty::KeyValue {
            key: Expression::string("java"),
            value: java },
    ]))
}

fn walk_callable_ref(pair: Pair<Rule>) -> Expression {
    let mut qualifier = None;
    let mut name = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::dotted_name => qualifier = Some(inner.as_str().to_string()),
            Rule::identifier | Rule::class_kw => name = Some(inner.as_str().to_string()),
            _ => {}
        }
    }

    let Some(name) = name else {
        return Expression::null();
    };
    if name == "class" {
        if let Some(qualifier) = qualifier {
            let is_type = qualifier
                .rsplit('.')
                .next()
                .and_then(|leaf| leaf.chars().next())
                .is_some_and(char::is_uppercase);
            return if is_type {
                kotlin_class_literal_expr(&qualifier)
            } else {
                Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__kt_class_of")),
                    args: vec![Argument::positional(dotted_ident_expr(&qualifier))],
                    optional: false })
            };
        }
        return Expression::null();
    }
    // `Int::plus`-style OPERATOR references are 2-parameter lambdas over the
    // binary op — the unbound-member form below would drop the second
    // operand.
    if let Some(ref qualifier) = qualifier
        && matches!(
            qualifier.rsplit('.').next().unwrap_or(qualifier),
            "Int" | "Long" | "Double" | "Float" | "String"
        )
        && let Some(op) = match name.as_str() {
            "plus" => Some(BinOp::Add),
            "minus" => Some(BinOp::Sub),
            "times" => Some(BinOp::Mul),
            "div" => Some(BinOp::Div),
            "rem" => Some(BinOp::Mod),
            _ => None }
    {
        let a = "__kt_ref_a";
        let b = "__kt_ref_b";
        return Expression::new(ExprKind::Lambda {
            params: vec![kt_param(a), kt_param(b)],
            body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Binary {
                op,
                left: Box::new(Expression::ident(a)),
                right: Box::new(Expression::ident(b)) }))),
            is_async: false,
            captures: Vec::new() });
    }
    // `Type::member` — an UNBOUND member reference: the lambda's argument is
    // the receiver (`String::length` is `{ it -> it.length }`), never a call
    // on the type object (`String.length(it)` dispatched nothing).
    if let Some(ref qualifier) = qualifier
        && matches!(
            qualifier.rsplit('.').next().unwrap_or(qualifier),
            "String" | "Int" | "Long" | "Double" | "Float" | "Char" | "Boolean"
                | "List" | "MutableList" | "Set" | "MutableSet" | "Map" | "MutableMap"
                | "CharSequence" | "Any" | "Pair" | "Triple"
        )
    {
        let arg_name = "__kt_ref_arg".to_string();
        let member = Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident(&arg_name)),
            field: name.clone(),
            null_safe: false });
        // Property spellings stay reads; anything else is a method call.
        let body = if matches!(
            name.as_str(),
            "length" | "size" | "first" | "second" | "third" | "keys" | "values"
                | "entries" | "indices" | "lastIndex"
        ) {
            member
        } else {
            Expression::new(ExprKind::Call {
                callee: Box::new(member),
                args: Vec::new(),
                optional: false })
        };
        return Expression::new(ExprKind::Lambda {
            params: vec![Param {
                name: arg_name,
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false }],
            body: LambdaBody::Expr(Box::new(body)),
            is_async: false,
            captures: Vec::new() });
    }
    // `::fn` UNQUALIFIED: whether the target is a user function (whose
    // VALUE keeps every arity and default) or a builtin needing a wrapper
    // lambda is only knowable after the whole module is walked — leave a
    // marker for the post-pass.
    if qualifier.is_none() {
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__kt_callable_ref")),
            args: vec![Argument::positional(Expression::new(ExprKind::Lit(
                Literal::Str(name.clone().into()),
            )))],
            optional: false });
    }
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
                alias }],
            level: 0 }
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
        span: Span::default() })
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
            if let ExprKind::Call { callee, args, .. } = &expr.kind {
                if matches!(&callee.kind, ExprKind::Ident(name) if name == "error") {
                    return Some(kotlin_error_throw_stmt(args));
                }
            }
            Some(repeat_to_for_in(&expr).unwrap_or_else(|| Statement::new(StmtKind::Expr(expr))))
        }
        Rule::expr => {
            let expr = walk_expr(inner_pair);
            Some(Statement::new(StmtKind::Expr(expr)))
        }
        _ => None };

    match (stmt, label_name) {
        (Some(s), Some(lbl)) => Some(Statement::new(StmtKind::Labeled {
            label: lbl,
            body: Box::new(s) })),
        (other, _) => other }
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
        LambdaBody::Expr(e) => vec![Statement::new(StmtKind::Expr((**e).clone()))] };
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
            optional: false }),
        body,
        of: true,
        else_body: None,
        is_async: false }))
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
                                    adjustments: vec![] }));
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
                                                        type_hint: decl.type_hint.as_deref().map(str::to_string),
                                                        init: decl.init,
                                                        modifiers: Modifiers {
                                                            visibility: Visibility::Public,
                                                            is_readonly: kind == VarDeclKind::Const,
                                                            ..Default::default()
                                                        },
                                                        with_events: false,
                                                        array_bounds: None });
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
        decorators }))
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
                        constructor_args: ctor_args });
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

    // A Kotlin `enum class` IS `java.lang.Enum` — its constants are OBJECTS
    // carrying `name`/`ordinal`, and `Color.RED == Color.RED` is reference
    // identity. `StmtKind::EnumDecl` is the OTHER enum model: the TS-shaped
    // bidirectional int object (`primitives/enums.rs`), where a constant is a
    // bare ordinal. Kotlin used it, so `println(Color.RED)` printed `0` and
    // `.name` / `.ordinal` were `undefined` — measured before this changed.
    //
    // So install the same JDK surface Java does, from the same place, and
    // differ only where the languages genuinely differ: Kotlin spells
    // `name`/`ordinal` as PROPERTIES.
    let constants: Vec<vybe_platform_jvm::lang_enum::EnumConstant> = members
        .iter()
        .map(|m| vybe_platform_jvm::lang_enum::EnumConstant {
            name: m.name.clone(),
            ctor_args: m
                .constructor_args
                .iter()
                .cloned()
                .map(Argument::positional)
                .collect() })
        .collect();
    vybe_platform_jvm::lang_enum::install(
        &name,
        &constants,
        &mut body_members,
        vybe_platform_jvm::lang_enum::Accessors::Properties,
    );

    Some(Statement::new(StmtKind::ClassDecl {
        name,
        parents: vec![],
        interfaces: vec![],
        members: body_members,
        modifiers: ClassModifiers::default(),
        decorators }))
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
                with_events: false }],
            kind: decl_kind.clone() })];

        for (idx, name) in names.into_iter().enumerate() {
            let read_expr = Expression::new(ExprKind::Index {
                object: Box::new(Expression::ident(&tmp_name)),
                index: Box::new(Expression::int(idx as i64)),
                null_safe: false });
            stmts.push(Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(name),
                    type_hint: None,
                    init: Some(read_expr),
                    array_bounds: None,
                    with_events: false }],
                kind: decl_kind.clone() }));
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
                with_events: false }],
            kind: if is_readonly {
                VarDeclKind::Const
            } else {
                VarDeclKind::Var
            } }))
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
                        Rule::type_ref => type_hint = kotlin_nullable_type_hint(csub.as_str()).0,
                        Rule::block => catch_block_stmts = walk_block_statements(csub),
                        _ => {}
                    }
                }
                // `catch (e: java.lang.StringIndexOutOfBoundsException)` —
                // thrown names are SIMPLE; strip the package qualifier.
                let simple = type_hint
                    .as_deref()
                    .map(|t| t.rsplit('.').next().unwrap_or(t).to_string());
                let types = match simple.as_deref() {
                    Some("Exception") | Some("Throwable") | None => vec![],
                    Some(t) => vec![t.to_string()] };
                catches.push(CatchClause {
                    types,
                    var_name: Some(param_name),
                    stack_var: None,
                    body: catch_block_stmts,
                    when_clause: None });
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
        finally }))
}

fn kotlin_block_statements_as_expr(mut stmts: Vec<Statement>) -> Expression {
    if stmts.len() == 1 {
        return match stmts.remove(0).kind {
            StmtKind::Expr(e) | StmtKind::Return(Some(e)) => e,
            _ => Expression::null() };
    }

    if let Some(last) = stmts.last_mut()
        && let StmtKind::Expr(expr) = std::mem::replace(&mut last.kind, StmtKind::Empty)
    {
        last.kind = StmtKind::Return(Some(expr));
    }

    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: Vec::new(),
            body: LambdaBody::Block(stmts),
            is_async: false,
            captures: Vec::new() })),
        args: Vec::new(),
        optional: false })
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
        // A DOTTED receiver (`Factory.Companion.`) is not the class: its
        // members are the companion's. Reading the CLASS members here bound
        // a `value` PARAMETER to `this.value`.
        .filter(|p| !p.as_str().trim_end_matches('.').contains('.'))
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
    let mut raw_param_types: Option<Vec<String>> = None;

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
                    return_type = kotlin_nullable_type_hint(inner.as_str()).0;
                }
            }
            Rule::identifier => {
                name = inner.as_str().to_string();
            }
            Rule::parameter_list => {
                let raw_tys: Vec<String> = inner
                    .clone()
                    .into_inner()
                    .filter(|p| p.as_rule() == Rule::parameter)
                    .map(|p| {
                        p.into_inner()
                            .find(|q| q.as_rule() == Rule::type_ref)
                            .map(|q| q.as_str().trim().to_string())
                            .unwrap_or_default()
                    })
                    .collect();
                raw_param_types = Some(raw_tys);
                params = walk_parameter_list(inner);
            }
            Rule::function_body_expr => {
                // `find`, not `next`: the `=` may sit at end-of-line with the
                // body on the next line, and the eaten NEWLINEs precede the
                // expr in the pair stream.
                if let Some(expr_pair) = inner.into_inner().find(|p| matches!(p.as_rule(), Rule::expr)) {
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
        // Builtin receivers have no CLASS_MEMBERS entry, so the walk-time
        // ext-receiver read resolved nothing — bind bare receiver members
        // (`second - first`, `keys`) onto `this` like `apply { }` bodies.
        let mut body_locals = HashSet::new();
        kotlin_local_binding_names(&body, &mut body_locals);
        for p in params.iter() {
            body_locals.insert(p.name.clone());
        }
        for stmt in &mut body {
            kotlin_rewrite_receiver_refs_stmt(stmt, "this", &body_locals);
        }
        let mut ext_params = vec![Param {
            name: "this".to_string(),
            type_hint: receiver_type.map(Into::into),
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false }];
        ext_params.extend(params);
        params = ext_params;
    }

    // Raw param type texts, in declaration order per name — generics only
    // survive HERE (the TypeHint spelling strips them), and same-arity
    // overload resolution needs `Pair<Int,…>` vs `Pair<String,…>`.
    if let Some(raw) = raw_param_types.take() {
        KOTLIN_FN_RAW_PARAM_TYPES.with(|map| {
            map.borrow_mut().entry(name.clone()).or_default().push(raw);
        });
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
        is_sub: false }))
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
                        let (hint, nullable) = kotlin_nullable_type_hint(p.as_str());
                        is_nullable = nullable;
                        type_hint = hint;
                    }
                    Rule::expr => default = Some(walk_expr(p)),
                    _ => {}
                }
            }
            let is_optional = default.is_some();
            // A `vararg` parameter's VALUE is an array — keeping the element
            // type as the hint made `v.size` dispatch through the STRING
            // length slot ("a,b,c".length = 5, not 3 elements).
            if is_rest {
                type_hint = Some("Array".to_string());
            }
            params.push(Param {
                name,
                type_hint: type_hint.map(Into::into),
                default,
                pass_by: PassBy::Value,
                is_rest,
                is_kwargs: false,
                is_optional,
                is_nullable });
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
            Rule::type_ref => type_hint = kotlin_nullable_type_hint(inner.as_str()).0,
            Rule::expr => init = Some(walk_expr(inner)),
            _ => {}
        }
    }

    if type_hint.is_none() {
        if let Some(ref expr) = init {
            match expr.kind {
                ExprKind::Array(_) => type_hint = Some("Array".to_string()),
                ExprKind::Object(_) => {
                    type_hint = kotlin_literal_keyed_collection_type(expr).map(str::to_string);
                }
                _ => {}
            }
        }
    }

    if !name.is_empty()
        && init
            .as_ref()
            .is_some_and(|expr| matches!(expr.kind, ExprKind::Tuple(_)))
    {
        KOTLIN_TUPLE_LOCALS.with(|set| {
            set.borrow_mut().insert(name.clone());
        });
    }

    Some(Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(name),
            type_hint: type_hint.map(Into::into),
            init,
            array_bounds: None,
            with_events: false }],
        kind: if is_const || is_readonly {
            VarDeclKind::Const
        } else {
            VarDeclKind::Var
        } }))
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
        _ => false }
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
    /// Class-local backing names for Kotlin properties that share their source
    /// spelling with a method. The shared class shape has one member table, so
    /// the frontend must keep the storage and callable names distinct.
    static KOTLIN_CLASS_FIELD_ALIASES: std::cell::RefCell<Vec<std::collections::HashMap<String, String>>> =
        std::cell::RefCell::new(Vec::new());
}

/// `field` inside an accessor is Kotlin's BACKING STORAGE, not a variable.
///
/// It is the only way a custom accessor can reach the property's own storage —
/// `set(v) { field = v + 1 }` — and reading it as a plain identifier left it
/// `undefined`. Every other identifier passes through untouched.
fn backing_field_substitution(name: &str) -> String {
    if name == "field" {
        return BACKING_FIELD.with(|stack| {
            stack
                .borrow()
                .last()
                .cloned()
                .unwrap_or_else(|| name.to_string())
        });
    }

    KOTLIN_CLASS_FIELD_ALIASES.with(|stack| {
        stack
            .borrow()
            .last()
            .and_then(|aliases| aliases.get(name).cloned())
            .unwrap_or_else(|| name.to_string())
    })
}

/// The storage name a property's `field` resolves to. Distinct per property, so
/// two properties in one class each writing `field` do not share one slot.
fn backing_field_name(property: &str) -> String {
    format!("__kt_field_{}", property)
}

fn class_body_method_names(pairs: &[Pair<Rule>]) -> std::collections::HashSet<String> {
    let mut names = std::collections::HashSet::new();
    for pair in pairs {
        if pair.as_rule() != Rule::class_body {
            continue;
        }
        for member_pair in pair.clone().into_inner() {
            if member_pair.as_rule() != Rule::class_member {
                continue;
            }
            let Some(inner_member) = member_pair.into_inner().next() else {
                continue;
            };
            if inner_member.as_rule() != Rule::function_decl {
                continue;
            }
            if let Some(id) = inner_member
                .into_inner()
                .find(|p| p.as_rule() == Rule::identifier)
            {
                names.insert(id.as_str().to_string());
            }
        }
    }
    names
}

fn class_property_method_collisions(
    pairs: &[Pair<Rule>],
    method_names: &std::collections::HashSet<String>,
) -> std::collections::HashMap<String, String> {
    let mut aliases = std::collections::HashMap::new();
    if method_names.is_empty() {
        return aliases;
    }

    for pair in pairs {
        match pair.as_rule() {
            Rule::primary_constructor => {
                for param in pair.clone().into_inner() {
                    if param.as_rule() != Rule::class_parameter {
                        continue;
                    }
                    let mut is_property = false;
                    let mut name = None;
                    for p in param.into_inner() {
                        match p.as_rule() {
                            Rule::val_kw | Rule::var_kw => is_property = true,
                            Rule::identifier => name = Some(p.as_str().to_string()),
                            _ => {}
                        }
                    }
                    if is_property {
                        if let Some(name) = name.filter(|n| method_names.contains(n)) {
                            aliases.insert(name.clone(), backing_field_name(&name));
                        }
                    }
                }
            }
            Rule::class_body => {
                for member_pair in pair.clone().into_inner() {
                    if member_pair.as_rule() != Rule::class_member {
                        continue;
                    }
                    let Some(inner_member) = member_pair.into_inner().next() else {
                        continue;
                    };
                    if inner_member.as_rule() != Rule::var_decl {
                        continue;
                    }
                    for p in inner_member.into_inner() {
                        if p.as_rule() == Rule::identifier {
                            let name = p.as_str().to_string();
                            if method_names.contains(&name) {
                                aliases.insert(name.clone(), backing_field_name(&name));
                            }
                            break;
                        }
                    }
                }
            }
            _ => {}
        }
    }

    aliases
}

fn with_kotlin_field_aliases<T>(
    aliases: &std::collections::HashMap<String, String>,
    f: impl FnOnce() -> T,
) -> T {
    if aliases.is_empty() {
        return f();
    }
    KOTLIN_CLASS_FIELD_ALIASES.with(|stack| stack.borrow_mut().push(aliases.clone()));
    let out = f();
    KOTLIN_CLASS_FIELD_ALIASES.with(|stack| {
        stack.borrow_mut().pop();
    });
    out
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
    let mut set_body: Vec<Statement> = Vec::new();
    let mut set_param = "value".to_string();
    for p in &inners {
        match p.as_rule() {
            Rule::property_accessor => {
                let mut is_get = false;
                for part in p.clone().into_inner() {
                    match part.as_rule() {
                        Rule::get_kw => is_get = true,
                        Rule::set_kw => is_get = false,
                        Rule::identifier => {
                            if !is_get {
                                set_param = part.as_str().to_string();
                            }
                        }
                        Rule::function_body_expr => {
                            if let Some(e) = part
                                .into_inner()
                                .find(|p| matches!(p.as_rule(), Rule::expr))
                            {
                                if is_get {
                                    body =
                                        vec![Statement::new(StmtKind::Return(Some(walk_expr(e))))];
                                } else {
                                    set_body =
                                        vec![Statement::new(StmtKind::Expr(walk_expr(e)))];
                                }
                            }
                        }
                        Rule::block => {
                            if is_get {
                                body = walk_block_statements(part.clone());
                            } else {
                                set_body = walk_block_statements(part.clone());
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

    if !set_body.is_empty() {
        let mut set_locals = HashSet::new();
        kotlin_local_binding_names(&set_body, &mut set_locals);
        set_locals.insert(set_param.clone());
        for stmt in &mut set_body {
            kotlin_rewrite_receiver_refs_stmt(stmt, "this", &set_locals);
        }
        let this_param = |hint: &str| Param {
            name: "this".to_string(),
            type_hint: Some(hint.into()),
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false };
        let mut value_param = this_param(&receiver);
        value_param.name = set_param;
        value_param.type_hint = None;
        PENDING_TOP_LEVEL_FNS.with(|pending| {
            pending.borrow_mut().push(Statement::new(StmtKind::FunctionDecl {
                name: format!("{name}__ext_set"),
                params: vec![this_param(&receiver), value_param],
                return_type: None,
                body: std::mem::take(&mut set_body),
                modifiers: Modifiers {
                    visibility: Visibility::Public,
                    is_extension: true,
                    ..Default::default()
                },
                handles: vec![],
                is_async: false,
                is_generator: false,
                is_sub: false }));
        });
    }

    // BUILTIN receivers (`Map`, `List`, `String`…) have no CLASS_MEMBERS
    // entry, so the ext-receiver read above resolves nothing — bind the
    // getter body's bare `keys`/`size`/method calls onto `this` the same
    // way `apply { }` bodies do.
    {
        let mut body_locals = HashSet::new();
        kotlin_local_binding_names(&body, &mut body_locals);
        for stmt in &mut body {
            kotlin_rewrite_receiver_refs_stmt(stmt, "this", &body_locals);
        }
    }

    Some(Statement::new(StmtKind::FunctionDecl {
        name,
        params: vec![Param {
            name: "this".to_string(),
            type_hint: Some(receiver.into()),
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false }],
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
        is_sub: false }))
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
            Rule::type_ref => type_hint = kotlin_nullable_type_hint(inner.as_str()).0,
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
                            if let Some(e) =
                                part.into_inner().find(|p| matches!(p.as_rule(), Rule::expr))
                            {
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
                            is_nullable: false },
                        body });
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
            array_bounds: None });
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
        } });
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

fn kt_this_member(field: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(Expression::new(ExprKind::This)),
        field: field.to_string(),
        null_safe: false })
}

fn kt_delegate_member(delegate_field: &str, field: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(kt_this_member(delegate_field)),
        field: field.to_string(),
        null_safe: false })
}

fn kt_return_member(expr: Expression) -> Vec<Statement> {
    vec![Statement::new(StmtKind::Return(Some(expr)))]
}

fn kt_delegate_property(delegate_field: &str, name: &str, readonly: bool) -> ClassMember {
    let setter = (!readonly).then(|| {
        let param = Param {
            name: "__kt_value".to_string(),
            type_hint: None,
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: true };
        PropertySetter {
            param,
            body: vec![Statement::new(StmtKind::Expr(Expression::new(
                ExprKind::Assign {
                    target: Box::new(kt_delegate_member(delegate_field, name)),
                    value: Box::new(Expression::ident("__kt_value")) },
            )))] }
    });
    ClassMember::Property {
        name: name.to_string(),
        type_hint: None,
        getter: Some(kt_return_member(kt_delegate_member(delegate_field, name))),
        setter,
        is_auto: false,
        modifiers: Modifiers {
            visibility: Visibility::Public,
            ..Default::default()
        } }
}

fn kt_delegate_method(name: &str, params: Vec<Param>, body: Vec<Statement>) -> ClassMember {
    ClassMember::Method(Box::new(Statement::new(StmtKind::FunctionDecl {
        name: name.to_string(),
        params,
        return_type: None,
        body,
        modifiers: Modifiers {
            visibility: Visibility::Public,
            ..Default::default()
        },
        handles: Vec::new(),
        is_async: false,
        is_generator: false,
        is_sub: false })))
}

fn kt_param(name: &str) -> Param {
    Param {
        name: name.to_string(),
        type_hint: None,
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: true }
}

fn kt_optional_param(name: &str, default: Expression) -> Param {
    Param {
        default: Some(default),
        is_optional: true,
        ..kt_param(name)
    }
}

fn kt_call(name: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(name)),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false })
}

fn member_exists(members: &[ClassMember], name: &str) -> bool {
    members.iter().any(|member| match member {
        ClassMember::Method(stmt) => {
            matches!(&stmt.kind, StmtKind::FunctionDecl { name: n, .. } if n == name)
        }
        ClassMember::Property { name: n, .. } | ClassMember::Field { name: n, .. } => n == name,
        _ => false })
}

fn append_kotlin_collection_delegate_members(
    members: &mut Vec<ClassMember>,
    interface_name: &str,
    delegate_field: &str,
) {
    let bare = interface_name.rsplit('.').next().unwrap_or(interface_name);
    let delegate = || kt_this_member(delegate_field);

    if matches!(bare, "List" | "MutableList" | "Collection" | "Iterable")
        && !member_exists(members, "size")
    {
        members.push(ClassMember::Property {
            name: "size".to_string(),
            type_hint: Some("Int".to_string()),
            getter: Some(kt_return_member(kt_call("__coll_length", vec![delegate()]))),
            setter: None,
            is_auto: false,
            modifiers: Modifiers {
                visibility: Visibility::Public,
                is_readonly: true,
                ..Default::default()
            } });
    }

    if matches!(bare, "List" | "MutableList") && !member_exists(members, "get") {
        members.push(kt_delegate_method(
            "get",
            vec![kt_param("index")],
            kt_return_member(Expression::new(ExprKind::Index {
                object: Box::new(delegate()),
                index: Box::new(Expression::ident("index")),
                null_safe: false })),
        ));
    }

    if matches!(bare, "List" | "MutableList" | "Collection") && !member_exists(members, "contains")
    {
        members.push(kt_delegate_method(
            "contains",
            vec![kt_param("value")],
            kt_return_member(kt_call(
                "__coll_contains",
                vec![delegate(), Expression::ident("value")],
            )),
        ));
    }

    if matches!(bare, "Set" | "MutableSet") && !member_exists(members, "contains") {
        members.push(kt_delegate_method(
            "contains",
            vec![kt_param("value")],
            kt_return_member(kotlin_set_contains_expr(
                delegate(),
                Expression::ident("value"),
            )),
        ));
    }

    if matches!(bare, "Map" | "MutableMap") {
        if !member_exists(members, "get") {
            members.push(kt_delegate_method(
                "get",
                vec![kt_param("key")],
                kt_return_member(kotlin_dict_get_expr(delegate(), Expression::ident("key"))),
            ));
        }
        if !member_exists(members, "containsKey") {
            members.push(kt_delegate_method(
                "containsKey",
                vec![kt_param("key")],
                kt_return_member(kotlin_dict_has_expr(delegate(), Expression::ident("key"))),
            ));
        }
        if !member_exists(members, "keys") {
            members.push(ClassMember::Property {
                name: "keys".to_string(),
                type_hint: None,
                getter: Some(kt_return_member(kt_call("__dict_keys", vec![delegate()]))),
                setter: None,
                is_auto: false,
                modifiers: Modifiers {
                    visibility: Visibility::Public,
                    is_readonly: true,
                    ..Default::default()
                } });
        }
    }

    if matches!(bare, "Set" | "MutableSet") && !member_exists(members, "size") {
        let raw_size = kt_call("__dict_size", vec![delegate()]);
        members.push(ClassMember::Property {
            name: "size".to_string(),
            type_hint: Some("Int".to_string()),
            getter: Some(kt_return_member(Expression::new(ExprKind::Binary {
                op: BinOp::Sub,
                left: Box::new(raw_size),
                right: Box::new(Expression::int(1)) }))),
            setter: None,
            is_auto: false,
            modifiers: Modifiers {
                visibility: Visibility::Public,
                is_readonly: true,
                ..Default::default()
            } });
    }

    if matches!(
        bare,
        "List" | "MutableList" | "Set" | "MutableSet" | "Collection" | "Iterable"
    ) && !member_exists(members, "joinToString")
    {
        members.push(kt_delegate_method(
            "joinToString",
            vec![kt_optional_param(
                "separator",
                Expression::new(ExprKind::Lit(Literal::Str(", ".to_string()))),
            )],
            kt_return_member(kt_call(
                "__coll_join",
                vec![delegate(), Expression::ident("separator")],
            )),
        ));
    }
}

fn append_kotlin_delegate_members(
    members: &mut Vec<ClassMember>,
    delegations: &[(String, String)],
) {
    for (interface_name, delegate_field) in delegations {
        CLASS_PROPERTIES.with(|props| {
            if let Some(source_props) = props.borrow().get(interface_name) {
                for (prop, readonly) in source_props {
                    if !member_exists(members, prop) {
                        members.push(kt_delegate_property(delegate_field, prop, *readonly));
                    }
                }
            }
        });
        append_kotlin_collection_delegate_members(members, interface_name, delegate_field);
        if !member_exists(members, "toString") {
            members.push(kt_delegate_method(
                "toString",
                Vec::new(),
                kt_return_member(Expression::new(ExprKind::Call {
                    callee: Box::new(kt_delegate_member(delegate_field, "toString")),
                    args: Vec::new(),
                    optional: false })),
            ));
        }
    }
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
    let class_method_names = class_body_method_names(&inner_pairs);
    let field_aliases = class_property_method_collisions(&inner_pairs, &class_method_names);
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
    let mut constructor_has_nullable_param = false;
    // Which member is the PRIMARY constructor, so property initializers and
    // `init` blocks land AFTER its parameter-to-field assignments.
    let mut primary_ctor_index: Option<usize> = None;
    // `: I by <expr>` — the field the forwarders read, and what fills it.
    // Kotlin stores the delegate; `AugmentDecl::via_field` names STORAGE, so a
    // `by` whose expression is not already a property needs one declared.
    let mut delegate_storage: Vec<(String, Expression)> = Vec::new();
    let mut delegations: Vec<(String, String)> = Vec::new();

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
                            Rule::type_ref => parent_name = type_hint_text(sub.as_str()),
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
                            delegations.push((parent_name.clone(), field.clone()));
                            // A delegating class IS the interface — `is I` and
                            // every interface-typed binding depend on it.
                            interfaces.push(parent_name.clone());
                            members.push(ClassMember::Augment(AugmentDecl {
                                from: parent_name.clone(),
                                via_field: Some(field),
                                adjustments: vec![] }));
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
                                adjustments: vec![] }));
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
                        let mut is_nullable = false;
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
                                Rule::type_ref => {
                                    let (hint, nullable) = kotlin_nullable_type_hint(p.as_str());
                                    is_nullable = nullable;
                                    constructor_has_nullable_param |= nullable;
                                    type_hint = hint;
                                }
                                Rule::expr => default = Some(walk_expr(p.clone())),
                                _ => {}
                            }
                        }
                        if !pname.is_empty() {
                            primary_prop_names.push(pname.clone());
                            let is_optional = default.is_some();
                            ctor_params.push(Param {
                                name: pname.clone(),
                                type_hint: type_hint.clone().map(Into::into),
                                default,
                                pass_by: PassBy::Value,
                                is_rest: false,
                                is_kwargs: false,
                                is_optional,
                                is_nullable });
                            if param_is_prop {
                                let field_name = field_aliases
                                    .get(&pname)
                                    .cloned()
                                    .unwrap_or_else(|| pname.clone());
                                members.push(ClassMember::Field {
                                    name: field_name.clone(),
                                    type_hint: type_hint.clone(),
                                    init: None,
                                    modifiers: Modifiers {
                                        visibility: Visibility::Public,
                                        is_readonly,
                                        ..Default::default()
                                    },
                                    with_events: false,
                                    array_bounds: None });
                                ctor_body.push(Statement::new(StmtKind::Expr(Expression::new(
                                    ExprKind::Assign {
                                        target: Box::new(Expression::new(ExprKind::Member {
                                            object: Box::new(Expression::new(ExprKind::This)),
                                            field: field_name,
                                            null_safe: false })),
                                        value: Box::new(Expression::ident(&pname)) },
                                ))));
                                let prop_idx = (primary_prop_names.len() - 1) as i64;
                                ctor_body.push(Statement::new(StmtKind::Expr(Expression::new(
                                    ExprKind::Assign {
                                        target: Box::new(Expression::new(ExprKind::Index {
                                            object: Box::new(Expression::new(ExprKind::This)),
                                            index: Box::new(Expression::int(prop_idx)),
                                            null_safe: false })),
                                        value: Box::new(Expression::ident(&pname)) },
                                ))));
                            }
                        }
                    }
                }

                primary_ctor_index = Some(members.len());
                KOTLIN_CLASS_PRIMARY_CTORS.with(|ctors| {
                    ctors.borrow_mut().insert(name.clone(), ctor_params.clone());
                });
                members.push(ClassMember::Constructor {
                    name: None,
                    params: ctor_params,
                    body: ctor_body,
                    base_args: base_args.clone(),
                    initializer_target: ConstructorInitializerTarget::Base,
                    visibility: Visibility::Public });
            }
            Rule::class_body => {
                for member_pair in inner.into_inner() {
                    if member_pair.as_rule() == Rule::class_member {
                        if let Some(inner_member) = member_pair.into_inner().next() {
                            match inner_member.as_rule() {
                                Rule::init_block => {
                                    if let Some(block_pair) = inner_member.into_inner().next() {
                                        let stmts =
                                            with_kotlin_field_aliases(&field_aliases, || {
                                                walk_block_statements(block_pair)
                                            });
                                        init_stmts.extend(stmts);
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
                                            Rule::block => {
                                                s_body = with_kotlin_field_aliases(
                                                    &field_aliases,
                                                    || walk_block_statements(sc),
                                                )
                                            }
                                            _ => {}
                                        }
                                    }
                                    members.push(ClassMember::Constructor {
                                        name: None,
                                        params: s_params,
                                        body: s_body,
                                        base_args: s_base_args,
                                        initializer_target: s_target,
                                        visibility: Visibility::Public });
                                }
                                Rule::class_decl | Rule::object_decl | Rule::interface_decl => {
                                    if let Some(stmt) = walk_statement(inner_member) {
                                        members.push(ClassMember::NestedType(Box::new(stmt)));
                                    }
                                }
                                Rule::function_decl => {
                                    if let Some(stmt) =
                                        with_kotlin_field_aliases(&field_aliases, || {
                                            walk_function_decl(inner_member)
                                        })
                                    {
                                        members.push(ClassMember::Method(Box::new(stmt)));
                                    }
                                }
                                Rule::var_decl => {
                                    // `val area: Int get() = w * h` is a
                                    // PROPERTY, not a field — it has no storage
                                    // and the accessor has to run on each read.
                                    let prop = with_kotlin_field_aliases(&field_aliases, || {
                                        walk_class_property(inner_member.clone())
                                    });
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
                                    if let Some(stmt) =
                                        with_kotlin_field_aliases(&field_aliases, || {
                                            walk_var_decl(inner_member)
                                        })
                                    {
                                        if let StmtKind::VarDecl { declarations, .. } = stmt.kind {
                                            for decl in declarations {
                                                if let BindingPattern::Ident(fname) = decl.pattern {
                                                    let field_name = field_aliases
                                                        .get(&fname)
                                                        .cloned()
                                                        .unwrap_or_else(|| fname.clone());
                                                    if is_const_val {
                                                        if let Some(val_expr) = decl.init {
                                                            members.push(ClassMember::Const {
                                                                name: field_name,
                                                                type_hint: decl.type_hint.as_deref().map(str::to_string),
                                                                value: val_expr,
                                                                visibility: Visibility::Public });
                                                        }
                                                    } else {
                                                        members.push(ClassMember::Field {
                                                            name: field_name,
                                                            type_hint: decl.type_hint.as_deref().map(str::to_string),
                                                            init: decl.init,
                                                            modifiers: Modifiers {
                                                                visibility: Visibility::Public,
                                                                is_readonly: is_readonly_val,
                                                                ..Default::default()
                                                            },
                                                            with_events: false,
                                                            array_bounds: None });
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

    append_kotlin_delegate_members(&mut members, &delegations);

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
                array_bounds: None });
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
                is_nullable: true })
            .collect();
        let capture_assigns: Vec<Statement> = captures
            .iter()
            .map(|cap| {
                Statement::new(StmtKind::Expr(Expression::new(ExprKind::Assign {
                    target: Box::new(Expression::new(ExprKind::Member {
                        object: Box::new(Expression::new(ExprKind::This)),
                        field: cap.clone(),
                        null_safe: false })),
                    value: Box::new(Expression::ident(cap)) })))
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
                visibility: Visibility::Public });
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
            array_bounds: None });
        let outer_param = Param {
            name: "__kt_outer".to_string(),
            type_hint: outer_name.map(Into::into),
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false };
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
                visibility: Visibility::Public });
        }
        for member in &mut members {
            if let ClassMember::Constructor { params, body, .. } = member {
                params.insert(0, outer_param.clone());
                body.insert(
                    0,
                    Statement::new(StmtKind::Expr(Expression::new(ExprKind::Assign {
                        target: Box::new(Expression::new(ExprKind::Member {
                            object: Box::new(Expression::new(ExprKind::This)),
                            field: "__kt_outer".to_string(),
                            null_safe: false })),
                        value: Box::new(Expression::ident("__kt_outer")) }))),
                );
            }
        }
    }

    if base_args.is_some()
        && !members
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
            visibility: Visibility::Public });
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
            array_bounds: None });
        if from_ctor_param {
            init_stmts.push(Statement::new(StmtKind::Expr(Expression::new(
                ExprKind::Assign {
                    target: Box::new(Expression::new(ExprKind::Member {
                        object: Box::new(Expression::new(ExprKind::This)),
                        field: field.clone(),
                        null_safe: false })),
                    value: Box::new(Expression::ident(field)) },
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
                visibility: Visibility::Public });
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

    if is_data {
        KOTLIN_DATA_CLASS_PROPERTY_INDEX.with(|map| {
            map.borrow_mut().insert(
                name.clone(),
                primary_prop_names
                    .iter()
                    .enumerate()
                    .map(|(idx, prop)| (prop.clone(), idx))
                    .collect(),
            );
        });
    }
    for (interface_name, _) in &delegations {
        if matches!(
            interface_name.rsplit('.').next().unwrap_or(interface_name),
            "List"
                | "MutableList"
                | "Set"
                | "MutableSet"
                | "Map"
                | "MutableMap"
                | "Collection"
                | "Iterable"
        ) {
            KOTLIN_DELEGATED_COLLECTIONS.with(|map| {
                map.borrow_mut()
                    .insert(name.clone(), interface_name.clone());
            });
        }
    }
    if constructor_has_nullable_param {
        KOTLIN_NULLABLE_CTOR_CLASSES.with(|set| {
            set.borrow_mut().insert(name.clone());
        });
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
        decorators }))
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
                                type_hint: decl.type_hint.as_deref().map(str::to_string),
                                value,
                                visibility: Visibility::Public });
                        }
                        continue;
                    }
                    members.push(ClassMember::Field {
                        name: fname,
                        type_hint: decl.type_hint.as_deref().map(str::to_string),
                        init: decl.init,
                        modifiers: Modifiers {
                            visibility: Visibility::Public,
                            is_static: statics,
                            is_readonly: is_readonly_val,
                            ..Default::default()
                        },
                        with_events: false,
                        array_bounds: None });
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
        decorators: vec![] }))
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
        else_body }))
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
/// The runtime test behind Kotlin `is`/`!is`/`as?`/`when is` for `type_name`
/// (already lowercased). Our numbers are ONE f64 tier, so `is Int`/`is Long`
/// additionally require an integer VALUE — `2.5 is Int` is false in real
/// Kotlin, but a bare number probe accepted it. The subject is cloned; call
/// sites pass cheap subjects (locals, `when` temps).
// ── kotlin.Result ───────────────────────────────────────────────────────────
// A Result is `{__kt_ok, __kt_v, __kt_e}`. Every operation lowers to ONE
// wrapper lambda over the receiver with the user lambda's body INLINED
// (never called nested: captured-variable writes only survive one lambda
// level), and Result-returning shapes ride the `__kt_as_result` identity so
// chains stay typed.

/// A receiver-typed function hint — `String.() -> String`, `Int.(Int) -> Int`.
/// Only `.(`: the stored spelling may have lost the `>` of `->` to the
/// generics eraser, and no other Kotlin type contains the sequence.
fn kotlin_hint_is_receiver_fn(hint: &str) -> bool {
    hint.contains(".(")
}

fn kotlin_type_is_result(t: &str) -> bool {
    t == "Result" || t.starts_with("Result<")
}

fn kotlin_result_obj(ok: bool, value: Expression) -> Expression {
    let key = |s: &str| Expression::new(ExprKind::Lit(Literal::Str(s.into())));
    Expression::new(ExprKind::Object(vec![
        ObjectProperty::KeyValue {
            key: key("__kt_ok"),
            value: Expression::new(ExprKind::Lit(Literal::Bool(ok))) },
        ObjectProperty::KeyValue {
            key: key(if ok { "__kt_v" } else { "__kt_e" }),
            value },
    ]))
}

fn kotlin_result_prop(recv: Expression, field: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(recv),
        field: field.to_string(),
        null_safe: false })
}

fn kotlin_as_result(e: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("__kt_as_result")),
        args: vec![Argument::positional(e)],
        optional: false })
}

/// One wrapper lambda `(__kr) -> { body }` applied to `recv`.
fn kotlin_result_iife(recv: Expression, body: Vec<Statement>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: kotlin_local_capture_params(&["__kr".to_string()]),
            body: LambdaBody::Block(body),
            is_async: false,
            captures: Vec::new() })),
        args: vec![Argument::positional(recv)],
        optional: false })
}

/// The user lambda's body as statements: `val <param> = arg` first, the
/// TAIL expression fed through `wrap` into a Return. `wrap` receiving the
/// tail decides the produced value; a body with no tail returns
/// `wrap(null)`.
fn kotlin_result_inline_lambda(
    lambda: &Expression,
    arg: Expression,
    wrap: &dyn Fn(Expression) -> Expression,
) -> Vec<Statement> {
    let ExprKind::Lambda { params, body, .. } = &lambda.kind else {
        // Not a literal lambda (a function reference): call it.
        let call = Expression::new(ExprKind::Call {
            callee: Box::new(lambda.clone()),
            args: vec![Argument::positional(arg)],
            optional: false });
        return vec![Statement::new(StmtKind::Return(Some(wrap(call))))];
    };
    let param = params
        .first()
        .map(|p| p.name.clone())
        .unwrap_or_else(|| "it".to_string());
    let mut out = vec![Statement::new(StmtKind::VarDecl {
        kind: vybe_ast::VarDeclKind::Const,
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(param),
            type_hint: None,
            init: Some(arg),
            array_bounds: None,
            with_events: false }] })];
    match body {
        LambdaBody::Expr(e) => out.push(Statement::new(StmtKind::Return(Some(wrap(
            (**e).clone(),
        ))))),
        LambdaBody::Block(stmts) => {
            let mut stmts = stmts.clone();
            match stmts.pop() {
                Some(Statement { kind: StmtKind::Expr(e), .. })
                | Some(Statement { kind: StmtKind::Return(Some(e)), .. }) => {
                    out.extend(stmts);
                    out.push(Statement::new(StmtKind::Return(Some(wrap(e)))));
                }
                Some(last @ Statement { kind: StmtKind::If { .. }, .. }) => {
                    out.extend(stmts);
                    let value = kotlin_if_stmt_to_ternary(last);
                    out.push(Statement::new(StmtKind::Return(Some(wrap(value)))));
                }
                Some(last) => {
                    out.extend(stmts);
                    out.push(last);
                    out.push(Statement::new(StmtKind::Return(Some(wrap(
                        Expression::null(),
                    )))));
                }
                None => out.push(Statement::new(StmtKind::Return(Some(wrap(
                    Expression::null(),
                ))))) }
        }
    }
    out
}

/// `if (__kr.__kt_ok == want_ok) { <then> } ; <tail>` — the branch shape
/// every Result combinator shares.
fn kotlin_result_branch(
    want_ok: bool,
    then: Vec<Statement>,
    tail: Vec<Statement>,
) -> Vec<Statement> {
    let ok = kotlin_result_prop(Expression::ident("__kr"), "__kt_ok");
    let cond = if want_ok {
        ok
    } else {
        Expression::new(ExprKind::Unary { op: UnaryOp::Not, expr: Box::new(ok) })
    };
    let mut out = vec![Statement::new(StmtKind::If {
        cond,
        then_body: then,
        elifs: Vec::new(),
        else_body: None })];
    out.extend(tail);
    out
}

/// `trunc(ms * num / den)` factors for the `inWhole*` accessors.
fn kotlin_duration_whole_factor(field: &str) -> Option<(f64, f64)> {
    Some(match field {
        "inWholeNanoseconds" => (1_000_000.0, 1.0),
        "inWholeMicroseconds" => (1_000.0, 1.0),
        "inWholeMilliseconds" => (1.0, 1.0),
        "inWholeSeconds" => (1.0, 1_000.0),
        "inWholeMinutes" => (1.0, 60_000.0),
        "inWholeHours" => (1.0, 3_600_000.0),
        "inWholeDays" => (1.0, 86_400_000.0),
        _ => return None })
}

/// The same factors keyed by a `DurationUnit.X` argument.
fn kotlin_duration_unit_factor(arg: &Expression) -> Option<(f64, f64)> {
    let ExprKind::Member { object, field, .. } = &arg.kind else {
        return None;
    };
    if !matches!(&object.kind, ExprKind::Ident(n) if n == "DurationUnit") {
        return None;
    }
    Some(match field.as_str() {
        "NANOSECONDS" => (1_000_000.0, 1.0),
        "MICROSECONDS" => (1_000.0, 1.0),
        "MILLISECONDS" => (1.0, 1.0),
        "SECONDS" => (1.0, 1_000.0),
        "MINUTES" => (1.0, 60_000.0),
        "HOURS" => (1.0, 3_600_000.0),
        "DAYS" => (1.0, 86_400_000.0),
        _ => return None })
}

fn kotlin_is_type_test(expr: Expression, type_name: String) -> Expression {
    // `f is Function1<*, *>` / `is Function2<…>` — the FunctionN interfaces
    // are "is it callable" at runtime.
    if type_name
        .split('<')
        .next()
        .is_some_and(|base| {
            let base = base.trim();
            base.strip_prefix("function")
                .is_some_and(|rest| rest.is_empty() || rest.chars().all(|c| c.is_ascii_digit()))
        })
    {
        return Expression::new(ExprKind::IsType {
            expr: Box::new(expr),
            type_name: "function".to_string() });
    }
    let integral = type_name == "int" || type_name == "long";
    let base = Expression::new(ExprKind::IsType {
        expr: Box::new(expr.clone()),
        type_name,
    });
    if !integral {
        return base;
    }
    let whole = Expression::new(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Mod,
            left: Box::new(expr),
            right: Box::new(Expression::int(1)),
        })),
        right: Box::new(Expression::int(0)),
    });
    Expression::new(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(base),
        right: Box::new(whole),
    })
}

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
                let test = kotlin_is_type_test(
                    subject.clone(),
                    type_hint_text(sub.as_str()).to_lowercase(),
                );
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
                        right: Box::new(walk_expr(lo)) });
                    let upper = Expression::new(ExprKind::Binary {
                        op: BinOp::LtEq,
                        left: Box::new(subject.clone()),
                        right: Box::new(walk_expr(hi)) });
                    out = Some(Expression::new(ExprKind::Binary {
                        op: BinOp::And,
                        left: Box::new(lower),
                        right: Box::new(upper) }));
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
                        right: Box::new(walk_expr(rhs)) }));
                }
            }
            Rule::expr => {
                let value = walk_expr(sub);
                let membership = |negated: bool| {
                    let test = Expression::new(ExprKind::Binary {
                        op: BinOp::In,
                        left: Box::new(subject.clone()),
                        right: Box::new(value.clone()) });
                    if negated { not_expr(test) } else { test }
                };
                out = Some(match prefix {
                    Some(Rule::in_kw) => membership(false),
                    Some(Rule::not_in_kw) => membership(true),
                    _ => Expression::new(ExprKind::Binary {
                        op: BinOp::Eq,
                        left: Box::new(subject.clone()),
                        right: Box::new(value) }) });
            }
            _ => {}
        }
    }

    out
}

fn not_expr(expr: Expression) -> Expression {
    Expression::new(ExprKind::Unary {
        op: UnaryOp::Not,
        expr: Box::new(expr) })
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
                                        to: walk_expr(e2) });
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
                                        expr: walk_expr(e) });
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
                body: body_stmts });
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
        default }))
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
                null_safe: false });
            prepended_stmts.push(Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(name),
                    type_hint: None,
                    init: Some(read_expr),
                    array_bounds: None,
                    with_events: false }],
                kind: VarDeclKind::Const }));
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
                optional: false })),
            then: Box::new(iter_expr.clone()),
            else_: Box::new(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__dict_items")),
                args: vec![Argument::positional(iter_expr)],
                optional: false })) })
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
        is_async: false }))
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
        else_body: None }))
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
        until: false }))
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
                                    let (hint, nullable) = kotlin_nullable_type_hint(lsub.as_str());
                                    lambda_param_nullable = nullable;
                                    type_hint = hint;
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
                                is_nullable: false });
                            for (idx, dname) in destruct_names.into_iter().enumerate() {
                                prefix_stmts.push(Statement::new(StmtKind::VarDecl {
                                    declarations: vec![VarDeclarator {
                                        pattern: BindingPattern::Ident(dname),
                                        type_hint: None,
                                        init: Some(Expression::new(ExprKind::Index {
                                            object: Box::new(Expression::ident(&tmp_param)),
                                            index: Box::new(Expression::int(idx as i64)),
                                            null_safe: false })),
                                        array_bounds: None,
                                        with_events: false }],
                                    kind: VarDeclKind::Const }));
                            }
                        } else if !name.is_empty() {
                            params.push(Param {
                                name,
                                type_hint: type_hint.map(Into::into),
                                default: None,
                                pass_by: PassBy::Value,
                                is_rest: false,
                                is_kwargs: false,
                                is_optional: false,
                                is_nullable: lambda_param_nullable });
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
            is_nullable: false });
    }

    kotlin_return_wrap_last(&mut body);

    Expression::new(ExprKind::Lambda {
        params,
        body: LambdaBody::Block(body),
        captures: vec![],
        is_async: false })
}

/// Return-wrap a lambda body's trailing statement. A trailing `if`/`else`
/// parses as a STATEMENT, so the plain Expr wrap missed it and the lambda
/// answered undefined (`mapIndexedNotNull { i, v -> if (…) v else null }`
/// dropped everything). Each branch tail is wrapped recursively.
fn kotlin_return_wrap_last(stmts: &mut Vec<Statement>) {
    if let Some(last) = stmts.pop() {
        match last.kind {
            StmtKind::Expr(e) => {
                stmts.push(Statement::new(StmtKind::Return(Some(e))));
            }
            StmtKind::If {
                cond,
                mut then_body,
                mut elifs,
                mut else_body } => {
                kotlin_return_wrap_last(&mut then_body);
                for (_, branch) in &mut elifs {
                    kotlin_return_wrap_last(branch);
                }
                if let Some(branch) = &mut else_body {
                    kotlin_return_wrap_last(branch);
                }
                stmts.push(Statement::new(StmtKind::If {
                    cond,
                    then_body,
                    elifs,
                    else_body }));
            }
            other => {
                stmts.push(Statement::new(other));
            }
        }
    }
}

fn return_last_expression(stmts: &mut Vec<Statement>) {
    if let Some(last) = stmts.pop() {
        match last.kind {
            StmtKind::Expr(expr) => stmts.push(Statement::new(StmtKind::Return(Some(expr)))),
            other => stmts.push(Statement::new(other)) }
    }
}

fn walk_try_expr(pair: Pair<Rule>) -> Expression {
    let Some(mut stmt) = walk_try_stmt(pair) else {
        return Expression::null();
    };
    if let StmtKind::Try { body, catches, .. } = &mut stmt.kind {
        return_last_expression(body);
        for catch in catches {
            return_last_expression(&mut catch.body);
        }
    }
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: Vec::new(),
            body: LambdaBody::Block(vec![stmt]),
            captures: Vec::new(),
            is_async: false })),
        args: Vec::new(),
        optional: false })
}

fn kotlin_exception_new_expr(exc_name: &str, message: Option<Expression>) -> Expression {
    if exc_name == "IllegalStateException" {
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__kt_illegal_state_exception")),
            args: message
                .map(Argument::positional)
                .into_iter()
                .collect::<Vec<_>>(),
            optional: false });
    }

    Expression::new(ExprKind::New {
        class: Box::new(Expression::ident(exc_name)),
        args: message
            .map(Argument::positional)
            .into_iter()
            .collect::<Vec<_>>() })
}

fn kotlin_error_throw_stmt(args: &[Argument]) -> Statement {
    let message = args.first().map(|arg| arg.value.clone());
    Statement::new(StmtKind::Throw {
        expr: Some(kotlin_exception_new_expr("IllegalStateException", message)),
        cause: None })
}

fn kotlin_error_throw_expr(args: &[Argument]) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: Vec::new(),
            body: LambdaBody::Block(vec![kotlin_error_throw_stmt(args)]),
            captures: Vec::new(),
            is_async: false })),
        args: Vec::new(),
        optional: false })
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
                        value: Box::new(rhs) })
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
                        _ => BinOp::Add };
                    Expression::new(ExprKind::Assign {
                        target: Box::new(lhs.clone()),
                        value: Box::new(Expression::new(ExprKind::Binary {
                            op: bin_op,
                            left: Box::new(lhs),
                            right: Box::new(rhs) })) })
                }
            } else {
                walk_expr(first)
            }
        }
        Rule::elvis => {
            let mut inner = pair.into_inner();
            let mut current = walk_expr(inner.next().unwrap());
            while let Some(_op) = inner.next() {
                let rhs_pair = inner.next().unwrap();
                let next_expr = if rhs_pair.as_rule() == Rule::elvis_escape {
                    // `v ?: return x` / `v ?: throw e` — control-flow RHS.
                    // Throw becomes a THROWING helper call (an expression);
                    // return becomes a MARKER that the statement pass lowers
                    // to a guard (`if (v == null) return x`).
                    let mut kw = None;
                    let mut payload = None;
                    for p in rhs_pair.into_inner() {
                        match p.as_rule() {
                            Rule::return_kw | Rule::throw_kw | Rule::break_kw
                            | Rule::continue_kw => kw = Some(p.as_rule()),
                            Rule::expr => payload = Some(walk_expr(p)),
                            _ => {}
                        }
                    }
                    let target = match kw {
                        Some(Rule::throw_kw) => "__kt_throw",
                        _ => "__kt_elvis_return" };
                    Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident(target)),
                        args: payload.into_iter().map(Argument::positional).collect(),
                        optional: false })
                } else {
                    walk_expr(rhs_pair)
                };
                current = Expression::new(ExprKind::NullCoalesce {
                    left: Box::new(current),
                    right: Box::new(next_expr) });
            }
            current
        }
        Rule::try_expr => walk_try_expr(pair),
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
                    _ => BinOp::Eq };
                current = Expression::new(ExprKind::Binary {
                    op,
                    left: Box::new(current),
                    right: Box::new(next_expr) });
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
                        right: Box::new(walk_expr(next_pair)) }),
                    "<=" => Expression::new(ExprKind::Binary {
                        op: BinOp::LtEq,
                        left: Box::new(current),
                        right: Box::new(walk_expr(next_pair)) }),
                    ">" => Expression::new(ExprKind::Binary {
                        op: BinOp::Gt,
                        left: Box::new(current),
                        right: Box::new(walk_expr(next_pair)) }),
                    ">=" => Expression::new(ExprKind::Binary {
                        op: BinOp::GtEq,
                        left: Box::new(current),
                        right: Box::new(walk_expr(next_pair)) }),
                    "in" => Expression::new(ExprKind::Binary {
                        op: BinOp::In,
                        left: Box::new(current),
                        right: Box::new(walk_expr(next_pair)) }),
                    "!in" => Expression::new(ExprKind::Unary {
                        op: UnaryOp::Not,
                        expr: Box::new(Expression::new(ExprKind::Binary {
                            op: BinOp::In,
                            left: Box::new(current),
                            right: Box::new(walk_expr(next_pair)) })) }),
                    "is" => kotlin_is_type_test(current, type_str),
                    "!is" => Expression::new(ExprKind::Unary {
                        op: UnaryOp::Not,
                        expr: Box::new(kotlin_is_type_test(current, type_str)) }),
                    _ => Expression::new(ExprKind::Binary {
                        op: BinOp::Eq,
                        left: Box::new(current),
                        right: Box::new(walk_expr(next_pair)) }) };
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
                    inclusive: true })
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
                    _ => BinOp::Add };
                current = Expression::new(ExprKind::Binary {
                    op,
                    left: Box::new(current),
                    right: Box::new(next_expr) });
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
                    _ => BinOp::Mul };
                current = Expression::new(ExprKind::Binary {
                    op,
                    left: Box::new(current),
                    right: Box::new(next_expr) });
            }
            current
        }
        Rule::type_cast => {
            let mut inner = pair.into_inner();
            let mut current = walk_expr(inner.next().unwrap());
            while let Some(op_pair) = inner.next() {
                let target_type = inner.next().unwrap().as_str().to_string();
                if op_pair.as_str().starts_with("as?") {
                    // `x as? T` — the SAFE cast answers null on mismatch,
                    // never a coerced value (`"abc" as? Int` was NaN).
                    // Same runtime test the `is` operator uses.
                    current = Expression::new(ExprKind::Ternary {
                        cond: Box::new(kotlin_is_type_test(
                            current.clone(),
                            type_hint_text(&target_type).to_lowercase(),
                        )),
                        then: Box::new(current),
                        else_: Box::new(Expression::null()) });
                } else {
                    current = Expression::new(ExprKind::Cast {
                        expr: Box::new(current),
                        type_name: target_type });
                }
            }
            current
        }
        Rule::infix_expr | Rule::infix_lvl => {
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
                        inclusive: false });
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
                                right: Box::new(Expression::int(1)) })),
                            Argument::positional(Expression::new(ExprKind::Unary {
                                op: UnaryOp::Neg,
                                expr: Box::new(Expression::int(1)) })),
                        ],
                        optional: false });
                } else if op_str == "step" {
                    // `range step n` — must convert range to 3-arg stepped form.
                    match current.kind.clone() {
                        // (a downTo b) step n  → replace -1 with -n
                        ExprKind::Call {
                            callee,
                            mut args,
                            optional } if matches!(&callee.kind, ExprKind::Ident(nm) if nm == "__kt_step_desc") =>
                        {
                            if args.len() == 3 {
                                args[2] = Argument::positional(Expression::new(ExprKind::Unary {
                                    op: UnaryOp::Neg,
                                    expr: Box::new(next_expr) }));
                            }
                            current = Expression::new(ExprKind::Call {
                                callee,
                                args,
                                optional });
                        }
                        // (a..b) step n  or  (a until b) step n
                        ExprKind::Range {
                            start,
                            end,
                            inclusive } => {
                            let stop = if inclusive {
                                // inclusive end+1 so the 3-arg exclusive loop includes end
                                Expression::new(ExprKind::Binary {
                                    op: BinOp::Add,
                                    left: end,
                                    right: Box::new(Expression::int(1)) })
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
                                optional: false });
                        }
                        _ => {
                            // Fallback: pass through as method call
                            current = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::new(ExprKind::Member {
                                    object: Box::new(current),
                                    field: "step".to_string(),
                                    null_safe: false })),
                                args: vec![Argument::positional(next_expr)],
                                optional: false });
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
                        right: Box::new(next_expr) });
                } else if is_extension_function(op_str) {
                    // A USER `infix fun X.word(y)` — extensions are free
                    // functions taking the receiver first; the member-call
                    // fallback below dispatched into nothing.
                    current = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident(op_str)),
                        args: vec![
                            Argument::positional(current),
                            Argument::positional(next_expr),
                        ],
                        optional: false });
                } else {
                    current = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Member {
                            object: Box::new(current),
                            field: op_str.to_string(),
                            null_safe: false })),
                        args: vec![Argument::positional(next_expr)],
                        optional: false });
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
                            "++" => UnaryOp::PreInc,
                            "--" => UnaryOp::PreDec,
                            _ => UnaryOp::Not };
                        current = Expression::new(ExprKind::Unary {
                            op: un_op,
                            expr: Box::new(current) });
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
                                                spread: is_spread });
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
                            null_safe } = current.clone().kind
                        {
                            if let Some(replacement) =
                                kotlin_lambda_tail_rewrite(object, field, &args)
                            {
                                current = replacement;
                                continue;
                            }
                            // General `x?.method(args)` — guard the receiver,
                            // then dispatch the PLAIN member call (the
                            // null-safe member lookup answered undefined for
                            // value-method receivers like strings/numbers).
                            // Skip every spelling the match below rewrites —
                            // those arms produce Index/builtin forms, and
                            // wrapping them as calls made `?.get(1)` dispatch
                            // an undefined member.
                            if null_safe
                                && !matches!(
                                    field.as_str(),
                                    "plus" | "minus" | "times" | "div" | "rem"
                                        | "put" | "get" | "getOrDefault" | "containsKey"
                                        | "startsWith" | "endsWith" | "find" | "findLast"
                                        | "contains" | "remove" | "clear" | "isEmpty"
                                        | "isNotEmpty" | "invoke"
                                )
                            {
                                current = Expression::new(ExprKind::Ternary {
                                    cond: Box::new(Expression::new(ExprKind::Binary {
                                        op: BinOp::Eq,
                                        left: object.clone(),
                                        right: Box::new(Expression::new(ExprKind::Lit(
                                            Literal::Null,
                                        ))) })),
                                    then: Box::new(Expression::new(ExprKind::Lit(
                                        Literal::Null,
                                    ))),
                                    else_: Box::new(Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::new(ExprKind::Member {
                                            object: object.clone(),
                                            field: field.clone(),
                                            null_safe: false })),
                                        args: args.clone(),
                                        optional: false })) });
                                continue;
                            }
                            // `(a..b).toString()` is `"a..b"` — the range
                            // itself is a materialized ARRAY here, which
                            // rendered `[a, …]`. `indices` produces a Range
                            // node, so this covers `arr.indices.toString()`.
                            if field == "toString"
                                && args.is_empty()
                                && let ExprKind::Range {
                                    ref start,
                                    ref end,
                                    inclusive } = object.kind
                            {
                                let shown_end = if inclusive {
                                    (**end).clone()
                                } else {
                                    Expression::new(ExprKind::Binary {
                                        op: BinOp::Sub,
                                        left: end.clone(),
                                        right: Box::new(Expression::int(1)) })
                                };
                                current = Expression::new(ExprKind::Binary {
                                    op: BinOp::Add,
                                    left: Box::new(Expression::new(ExprKind::Binary {
                                        op: BinOp::Add,
                                        left: start.clone(),
                                        right: Box::new(Expression::new(ExprKind::Lit(
                                            Literal::Str("..".into()),
                                        ))) })),
                                    right: Box::new(shown_end) });
                                continue;
                            }
                            match field.as_str() {
                                // Operator functions spelled as calls:
                                // `a.plus(b)` is `a + b`; `a?.plus(b)` guards
                                // the receiver first (a bare member lookup on
                                // a NUMBER found nothing callable).
                                "plus" | "minus" | "times" | "div" | "rem"
                                    if args.len() == 1 =>
                                {
                                    let op = match field.as_str() {
                                        "plus" => BinOp::Add,
                                        "minus" => BinOp::Sub,
                                        "times" => BinOp::Mul,
                                        "div" => BinOp::Div,
                                        _ => BinOp::Mod };
                                    let bin = Expression::new(ExprKind::Binary {
                                        op,
                                        left: object.clone(),
                                        right: Box::new(args[0].value.clone()) });
                                    current = if null_safe {
                                        Expression::new(ExprKind::Ternary {
                                            cond: Box::new(Expression::new(ExprKind::Binary {
                                                op: BinOp::Eq,
                                                left: object.clone(),
                                                right: Box::new(Expression::new(ExprKind::Lit(
                                                    Literal::Null,
                                                ))) })),
                                            then: Box::new(Expression::new(ExprKind::Lit(
                                                Literal::Null,
                                            ))),
                                            else_: Box::new(bin) })
                                    } else {
                                        bin
                                    };
                                    continue;
                                }
                                // `f.invoke(args)` IS `f(args)` — the function
                                // value has no member table. `f?.invoke(…)`
                                // guards the callee first.
                                "invoke" => {
                                    let direct = Expression::new(ExprKind::Call {
                                        callee: object.clone(),
                                        args: args.clone(),
                                        optional: false });
                                    current = if null_safe {
                                        Expression::new(ExprKind::Ternary {
                                            cond: Box::new(Expression::new(ExprKind::Binary {
                                                op: BinOp::Eq,
                                                left: object.clone(),
                                                right: Box::new(Expression::new(ExprKind::Lit(
                                                    Literal::Null,
                                                ))) })),
                                            then: Box::new(Expression::new(ExprKind::Lit(
                                                Literal::Null,
                                            ))),
                                            else_: Box::new(direct) })
                                    } else {
                                        direct
                                    };
                                    continue;
                                }
                                "put" if args.len() == 2 => {
                                    // Kotlin's `put` RETURNS the previous value
                                    // (or null) — a bare index-assign answers the
                                    // assigned value instead.
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__kt_map_put")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            Argument::positional(args[0].value.clone()),
                                            Argument::positional(args[1].value.clone()),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "get" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Index {
                                        object: object.clone(),
                                        index: Box::new(args[0].value.clone()),
                                        null_safe: false });
                                    continue;
                                }
                                "getOrDefault" if args.len() == 2 => {
                                    current = Expression::new(ExprKind::NullCoalesce {
                                        left: Box::new(Expression::new(ExprKind::Index {
                                            object: object.clone(),
                                            index: Box::new(args[0].value.clone()),
                                            null_safe: false })),
                                        right: Box::new(args[1].value.clone()) });
                                    continue;
                                }
                                "containsKey" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Binary {
                                        op: BinOp::In,
                                        left: Box::new(args[0].value.clone()),
                                        right: object.clone() });
                                    continue;
                                }
                                "startsWith" | "endsWith" if args.len() == 1 => {
                                    // Free-call form of the [builtins] entry —
                                    // the member path missed for literal
                                    // receivers (measured: `"a".startsWith("a")`
                                    // was "undefined is not callable").
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident(field)),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "startsWith" | "endsWith" if args.len() == 1 => {
                                    // Free-call form of the [builtins] entry —
                                    // the member path missed for literal
                                    // receivers (measured: `"a".startsWith("a")`
                                    // was "undefined is not callable").
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident(field)),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "startsWith" | "endsWith" if args.len() == 2 => {
                                    let target = if field == "startsWith" {
                                        "__kt_starts_with_ic"
                                    } else {
                                        "__kt_ends_with_ic"
                                    };
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident(target)),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                            args[1].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "startsWith" | "endsWith" if args.len() == 1 => {
                                    // Free-call form of the [builtins] entry —
                                    // the member path missed for literal
                                    // receivers (measured: `"a".startsWith("a")`
                                    // was "undefined is not callable").
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident(field)),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "startsWith" | "endsWith" if args.len() == 2 => {
                                    let target = if field == "startsWith" {
                                        "__kt_starts_with_ic"
                                    } else {
                                        "__kt_ends_with_ic"
                                    };
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident(target)),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                            args[1].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "find" | "findLast" if args.len() == 1 => {
                                    let target = if field == "find" {
                                        "__kt_find"
                                    } else {
                                        "__kt_find_last"
                                    };
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident(target)),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "contains" if args.len() == 2 => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__kt_contains_ic")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                            args[1].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "startsWith" | "endsWith" if args.len() == 1 => {
                                    // Free-call form of the [builtins] entry —
                                    // the member path missed for literal
                                    // receivers (measured: `"a".startsWith("a")`
                                    // was "undefined is not callable").
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident(field)),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "startsWith" | "endsWith" if args.len() == 2 => {
                                    let target = if field == "startsWith" {
                                        "__kt_starts_with_ic"
                                    } else {
                                        "__kt_ends_with_ic"
                                    };
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident(target)),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                            args[1].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "find" | "findLast" if args.len() == 1 => {
                                    let target = if field == "find" {
                                        "__kt_find"
                                    } else {
                                        "__kt_find_last"
                                    };
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident(target)),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "contains" if args.len() == 2 => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__kt_contains_ic")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                            args[1].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "contains" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_contains")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            Argument::positional(args[0].value.clone()),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                // NOTE: `.add(x)` for Set (dict) is handled in the second
                                // Member block below via __coll_push, which works uniformly
                                // for both list (array.push) and set (set semantics via
                                // array.push on the keys array). Do NOT intercept it here.
                                "remove" if args.len() == 1 || args.len() == 2 => {
                                    // 1-arg: remove by key (map) / by value (list),
                                    // answering the removed value / a Boolean.
                                    // 2-arg: `remove(key, value)` — conditional
                                    // remove, answering a Boolean. The argument
                                    // stays RAW — `list.remove(3)` removes the
                                    // VALUE 3, and a stringified "3" matches
                                    // nothing; the adapter stringifies on its
                                    // map branch only.
                                    let mut call_args = vec![
                                        Argument::positional(*object.clone()),
                                        Argument::positional(args[0].value.clone()),
                                    ];
                                    if args.len() == 2 {
                                        call_args
                                            .push(Argument::positional(args[1].value.clone()));
                                    }
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__dict_delete")),
                                        args: call_args,
                                        optional: false });
                                    continue;
                                }
                                "clear" if args.is_empty() => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__dict_clear")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false });
                                    continue;
                                }
                                // No-arg `any()`/`none()` are EMPTINESS tests —
                                // `[array_methods] any` maps to `some`, which
                                // calls a callback that is not there.
                                "any" if args.is_empty() => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__kt_is_not_empty")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false });
                                    continue;
                                }
                                "none" if args.is_empty() => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__kt_is_empty")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false });
                                    continue;
                                }
                                "isEmpty" if args.is_empty() => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__kt_is_empty")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false });
                                    continue;
                                }
                                "isNotEmpty" if args.is_empty() => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__kt_is_not_empty")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false });
                                    continue;
                                }
                                _ => {}
                            }
                        }

                        let func_name = match &current.kind {
                            ExprKind::Ident(name) => Some(name.clone()),
                            _ => None };

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
                            let builder_lambda = args.len() == 1
                                && matches!(args[0].value.kind, ExprKind::Lambda { .. });
                            if builder_lambda
                                && matches!(fn_name.as_str(), "buildMap" | "buildSet" | "buildList")
                            {
                                // `buildX { … }` is the builder-lambda form —
                                // post-alias lowering makes it an IIFE; the
                                // literal conversions below would swallow the
                                // lambda as a bogus element.
                            } else if matches!(
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
                                                value: v });
                                            continue;
                                        }
                                    }
                                    if let ExprKind::Tuple(ref pair_elems) = arg.value.kind {
                                        if pair_elems.len() == 2 {
                                            props.push(ObjectProperty::KeyValue {
                                                key: pair_elems[0].clone(),
                                                value: pair_elems[1].clone() });
                                            continue;
                                        }
                                    }
                                    if let ExprKind::Array(ref pair_elems) = arg.value.kind {
                                        if pair_elems.len() == 2 {
                                            props.push(ObjectProperty::KeyValue {
                                                key: pair_elems[0].value.clone(),
                                                value: pair_elems[1].value.clone() });
                                            continue;
                                        }
                                    }
                                    props.push(ObjectProperty::KeyValue {
                                        key: Expression::new(ExprKind::Index {
                                            object: Box::new(arg.value.clone()),
                                            index: Box::new(Expression::int(0)),
                                            null_safe: false }),
                                        value: Expression::new(ExprKind::Index {
                                            object: Box::new(arg.value.clone()),
                                            index: Box::new(Expression::int(1)),
                                            null_safe: false }) });
                                }
                                current = create_map_expr(props);
                                continue;
                            }
                            if !builder_lambda
                                && matches!(
                                    fn_name.as_str(),
                                    "setOf"
                                        | "mutableSetOf"
                                        | "linkedSetOf"
                                        | "hashSetOf"
                                        | "buildSet"
                                        | "emptySet"
                                )
                            {
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
                                    optional: false });
                                continue;
                            }
                            if matches!(
                                fn_name.as_str(),
                                "listOf"
                                    | "mutableListOf"
                                    | "arrayListOf"
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
                                        by_ref: false })
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
                                    optional: false });
                                continue;
                            }
                        }

                        if let ExprKind::Member { ref mut field, .. } = current.kind {
                            if let Some(storage_name) =
                                overloaded_storage_name_for_args(field, args.len(), &args)
                            {
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
                            _ => false };
                        if let ExprKind::Member {
                            ref object,
                            ref field,
                            ..
                        } = current.kind
                            && rewritable
                        {
                            if let Some(replacement) =
                                kotlin_lambda_tail_rewrite(object, field, &args)
                            {
                                current = replacement;
                                continue;
                            }
                            match field.as_str() {
                                // `f.invoke(args)` IS `f(args)` — the function
                                // value has no member table.
                                "invoke" => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: object.clone(),
                                        args: args.clone(),
                                        optional: false });
                                    continue;
                                }
                                // `w.transform(5)` where `transform` is a
                                // function-VALUED property: read the property,
                                // then call the value (method dispatch found
                                // no member).
                                prop if is_user_property_name(prop)
                                    && !is_user_member_name(prop, args.len()) =>
                                {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::new(ExprKind::Call {
                                            callee: Box::new(Expression::ident(
                                                "__kt_safe_get",
                                            )),
                                            args: vec![
                                                Argument::positional(*object.clone()),
                                                Argument::positional(Expression::new(
                                                    ExprKind::Lit(Literal::Str(
                                                        prop.to_string().into(),
                                                    )),
                                                )),
                                            ],
                                            optional: false })),
                                        args: args.clone(),
                                        optional: false });
                                    continue;
                                }
                                "put" if args.len() == 2 => {
                                    // Kotlin's `put` RETURNS the previous value
                                    // (or null) — a bare index-assign answers the
                                    // assigned value instead.
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__kt_map_put")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            Argument::positional(args[0].value.clone()),
                                            Argument::positional(args[1].value.clone()),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "get" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Index {
                                        object: object.clone(),
                                        index: Box::new(args[0].value.clone()),
                                        null_safe: false });
                                    continue;
                                }
                                "getOrDefault" if args.len() == 2 => {
                                    let get_expr = Expression::new(ExprKind::Index {
                                        object: object.clone(),
                                        index: Box::new(args[0].value.clone()),
                                        null_safe: false });
                                    current = Expression::new(ExprKind::NullCoalesce {
                                        left: Box::new(get_expr),
                                        right: Box::new(args[1].value.clone()) });
                                    continue;
                                }
                                "containsKey" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Binary {
                                        op: BinOp::In,
                                        left: Box::new(args[0].value.clone()),
                                        right: object.clone() });
                                    continue;
                                }
                                "startsWith" | "endsWith" if args.len() == 1 => {
                                    // Free-call form of the [builtins] entry —
                                    // the member path missed for literal
                                    // receivers (measured: `"a".startsWith("a")`
                                    // was "undefined is not callable").
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident(field)),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "startsWith" | "endsWith" if args.len() == 1 => {
                                    // Free-call form of the [builtins] entry —
                                    // the member path missed for literal
                                    // receivers (measured: `"a".startsWith("a")`
                                    // was "undefined is not callable").
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident(field)),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "startsWith" | "endsWith" if args.len() == 2 => {
                                    let target = if field == "startsWith" {
                                        "__kt_starts_with_ic"
                                    } else {
                                        "__kt_ends_with_ic"
                                    };
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident(target)),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                            args[1].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "startsWith" | "endsWith" if args.len() == 1 => {
                                    // Free-call form of the [builtins] entry —
                                    // the member path missed for literal
                                    // receivers (measured: `"a".startsWith("a")`
                                    // was "undefined is not callable").
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident(field)),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "startsWith" | "endsWith" if args.len() == 2 => {
                                    let target = if field == "startsWith" {
                                        "__kt_starts_with_ic"
                                    } else {
                                        "__kt_ends_with_ic"
                                    };
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident(target)),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                            args[1].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "find" | "findLast" if args.len() == 1 => {
                                    let target = if field == "find" {
                                        "__kt_find"
                                    } else {
                                        "__kt_find_last"
                                    };
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident(target)),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "contains" if args.len() == 2 => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__kt_contains_ic")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                            args[1].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "startsWith" | "endsWith" if args.len() == 1 => {
                                    // Free-call form of the [builtins] entry —
                                    // the member path missed for literal
                                    // receivers (measured: `"a".startsWith("a")`
                                    // was "undefined is not callable").
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident(field)),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "startsWith" | "endsWith" if args.len() == 2 => {
                                    let target = if field == "startsWith" {
                                        "__kt_starts_with_ic"
                                    } else {
                                        "__kt_ends_with_ic"
                                    };
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident(target)),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                            args[1].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "find" | "findLast" if args.len() == 1 => {
                                    let target = if field == "find" {
                                        "__kt_find"
                                    } else {
                                        "__kt_find_last"
                                    };
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident(target)),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "contains" if args.len() == 2 => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__kt_contains_ic")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                            args[1].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "contains" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_contains")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            Argument::positional(args[0].value.clone()),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "containsValue" if args.len() == 1 => {
                                    let values_expr = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__dict_values")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false });
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_contains")),
                                        args: vec![
                                            Argument::positional(values_expr),
                                            Argument::positional(args[0].value.clone()),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                // No-arg `any()`/`none()` are EMPTINESS tests —
                                // `[array_methods] any` maps to `some`, which
                                // calls a callback that is not there.
                                "any" if args.is_empty() => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__kt_is_not_empty")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false });
                                    continue;
                                }
                                "none" if args.is_empty() => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__kt_is_empty")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false });
                                    continue;
                                }
                                "isEmpty" if args.is_empty() => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__kt_is_empty")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false });
                                    continue;
                                }
                                "isNotEmpty" if args.is_empty() => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__kt_is_not_empty")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false });
                                    continue;
                                }
                                "remove" if args.len() == 1 || args.len() == 2 => {
                                    // 1-arg: remove by key (map) / by value (list),
                                    // answering the removed value / a Boolean.
                                    // 2-arg: `remove(key, value)` — conditional
                                    // remove, answering a Boolean. The argument
                                    // stays RAW — `list.remove(3)` removes the
                                    // VALUE 3, and a stringified "3" matches
                                    // nothing; the adapter stringifies on its
                                    // map branch only.
                                    let mut call_args = vec![
                                        Argument::positional(*object.clone()),
                                        Argument::positional(args[0].value.clone()),
                                    ];
                                    if args.len() == 2 {
                                        call_args
                                            .push(Argument::positional(args[1].value.clone()));
                                    }
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__dict_delete")),
                                        args: call_args,
                                        optional: false });
                                    continue;
                                }
                                // `sort(fromIndex, toIndex)` — the range form;
                                // `[array_methods] sort` would read the bounds
                                // as a comparator. NOT `Collections.sort(list,
                                // cmp)` — that is the STATIC 2-arg form.
                                "sort"
                                    if args.len() == 2
                                        && !matches!(
                                            object.kind,
                                            ExprKind::Ident(ref n)
                                                if n == "Collections" || n == "Arrays"
                                        )
                                        && !matches!(
                                            object.kind,
                                            ExprKind::Member { ref field, .. }
                                                if field == "Collections" || field == "Arrays"
                                        )
                                        && !args.iter().any(|a| {
                                            matches!(
                                                a.value.kind,
                                                ExprKind::Lambda { .. }
                                                    | ExprKind::Call { .. }
                                                    | ExprKind::New { .. }
                                            )
                                        }) =>
                                {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__kt_sort_range")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                            args[1].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "removeAt" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_removeAt")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            Argument::positional(args[0].value.clone()),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "clear" if args.is_empty() => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__dict_clear")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false });
                                    continue;
                                }
                                "add" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__kt_add")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            Argument::positional(args[0].value.clone()),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "addAll" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__kt_add_all")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            Argument::positional(args[0].value.clone()),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "add" if args.len() == 2 => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__jvm_add")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            Argument::positional(args[0].value.clone()),
                                            Argument::positional(args[1].value.clone()),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "indexOf" if args.len() == 2 => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__kt_index_of_from")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                            args[1].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "indexOf" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_indexOf")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            Argument::positional(args[0].value.clone()),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "lastIndexOf" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_lastIndexOf")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            Argument::positional(args[0].value.clone()),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "reversed" if args.is_empty() => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_reverse")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false });
                                    continue;
                                }
                                "sorted" if args.is_empty() => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_sorted")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false });
                                    continue;
                                }
                                "joinToString"
                                    if args.iter().any(|a| a.name.is_some()) =>
                                {
                                    // Named-argument form: `joinToString(prefix = "[",
                                    // postfix = "]", separator = ",")`. The positional
                                    // arms below read args[0] as the separator, which
                                    // called the prefix STRING as a function.
                                    let by_name = |wanted: &str| {
                                        args.iter()
                                            .find(|a| a.name.as_deref() == Some(wanted))
                                            .map(|a| a.value.clone())
                                    };
                                    let sep = by_name("separator")
                                        .or_else(|| {
                                            args.iter()
                                                .find(|a| a.name.is_none())
                                                .map(|a| a.value.clone())
                                        })
                                        .unwrap_or_else(|| {
                                            Expression::string(", ")
                                        });
                                    let joined = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_join")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            Argument::positional(sep),
                                        ],
                                        optional: false });
                                    let mut out = joined;
                                    if let Some(prefix) = by_name("prefix") {
                                        out = Expression::new(ExprKind::Binary {
                                            op: BinOp::Add,
                                            left: Box::new(prefix),
                                            right: Box::new(out) });
                                    }
                                    if let Some(postfix) = by_name("postfix") {
                                        out = Expression::new(ExprKind::Binary {
                                            op: BinOp::Add,
                                            left: Box::new(out),
                                            right: Box::new(postfix) });
                                    }
                                    current = out;
                                    continue;
                                }
                                "joinToString"
                                    if args.iter().any(|a| a.name.is_some()) =>
                                {
                                    // Named-argument form: `joinToString(prefix = "[",
                                    // postfix = "]", separator = ",")`. The positional
                                    // arms below read args[0] as the separator, which
                                    // called the prefix STRING as a function.
                                    let by_name = |wanted: &str| {
                                        args.iter()
                                            .find(|a| a.name.as_deref() == Some(wanted))
                                            .map(|a| a.value.clone())
                                    };
                                    let sep = by_name("separator")
                                        .or_else(|| {
                                            args.iter()
                                                .find(|a| a.name.is_none())
                                                .map(|a| a.value.clone())
                                        })
                                        .unwrap_or_else(|| {
                                            Expression::string(", ")
                                        });
                                    let joined = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_join")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            Argument::positional(sep),
                                        ],
                                        optional: false });
                                    let mut out = joined;
                                    if let Some(prefix) = by_name("prefix") {
                                        out = Expression::new(ExprKind::Binary {
                                            op: BinOp::Add,
                                            left: Box::new(prefix),
                                            right: Box::new(out) });
                                    }
                                    if let Some(postfix) = by_name("postfix") {
                                        out = Expression::new(ExprKind::Binary {
                                            op: BinOp::Add,
                                            left: Box::new(out),
                                            right: Box::new(postfix) });
                                    }
                                    current = out;
                                    continue;
                                }
                                "joinToString" if args.len() >= 2 => {
                                    if let Some(items) = kotlin_static_array_items(object) {
                                        if let Some(mapped) = kotlin_apply_static_join_transform(
                                            &items,
                                            &args[1].value,
                                        ) {
                                            current = Expression::new(ExprKind::Call {
                                                callee: Box::new(Expression::ident("__coll_join")),
                                                args: vec![
                                                    Argument::positional(mapped),
                                                    Argument::positional(args[0].value.clone()),
                                                ],
                                                optional: false });
                                            continue;
                                        }
                                    }
                                    let mapped =
                                        kotlin_map_call_expr(*object.clone(), args[1].clone());
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_join")),
                                        args: vec![
                                            Argument::positional(mapped),
                                            Argument::positional(args[0].value.clone()),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "joinToString" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_join")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            Argument::positional(args[0].value.clone()),
                                        ],
                                        optional: false });
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
                                        optional: false });
                                    continue;
                                }
                                // `sortedBy` is NOT rewritten to `sort`: its
                                // lambda is a 1-arg KEY SELECTOR, and `sort`'s
                                // is a 2-arg comparator. `[array_methods]`
                                // routes the surviving member call to
                                // `__array_sort_by_key`.
                                // `ignoreCase = true` (and `limit = n`) named
                                // flags become positional so builtin arity
                                // dispatch sees them.
                                "contains" | "startsWith" | "endsWith" | "equals"
                                | "replace" | "replaceFirst" | "indexOf" | "lastIndexOf"
                                | "split" | "regionMatches"
                                    if args.iter().any(|a| a.name.is_some()) =>
                                {
                                    let mut new_args = args.clone();
                                    for a in &mut new_args {
                                        a.name = None;
                                    }
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::new(ExprKind::Member {
                                            object: object.clone(),
                                            field: field.clone(),
                                            null_safe: false })),
                                        args: new_args,
                                        optional: false });
                                    continue;
                                }
                                // 2-arg `contains(needle, ignoreCase)` — the
                                // 1-arg form was rewritten above; this one
                                // must dodge `[array_methods]`' `contains`.
                                "startsWith" | "endsWith" if args.len() == 1 => {
                                    // Free-call form of the [builtins] entry —
                                    // the member path missed for literal
                                    // receivers (measured: `"a".startsWith("a")`
                                    // was "undefined is not callable").
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident(field)),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "startsWith" | "endsWith" if args.len() == 1 => {
                                    // Free-call form of the [builtins] entry —
                                    // the member path missed for literal
                                    // receivers (measured: `"a".startsWith("a")`
                                    // was "undefined is not callable").
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident(field)),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "startsWith" | "endsWith" if args.len() == 2 => {
                                    let target = if field == "startsWith" {
                                        "__kt_starts_with_ic"
                                    } else {
                                        "__kt_ends_with_ic"
                                    };
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident(target)),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                            args[1].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "startsWith" | "endsWith" if args.len() == 1 => {
                                    // Free-call form of the [builtins] entry —
                                    // the member path missed for literal
                                    // receivers (measured: `"a".startsWith("a")`
                                    // was "undefined is not callable").
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident(field)),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "startsWith" | "endsWith" if args.len() == 2 => {
                                    let target = if field == "startsWith" {
                                        "__kt_starts_with_ic"
                                    } else {
                                        "__kt_ends_with_ic"
                                    };
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident(target)),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                            args[1].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "find" | "findLast" if args.len() == 1 => {
                                    let target = if field == "find" {
                                        "__kt_find"
                                    } else {
                                        "__kt_find_last"
                                    };
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident(target)),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "contains" if args.len() == 2 => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__kt_contains_ic")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                            args[1].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                // `ignoreCase = true` (and `limit = n`) named
                                // flags become positional so builtin arity
                                // dispatch sees them.
                                "contains" | "startsWith" | "endsWith" | "equals"
                                | "replace" | "replaceFirst" | "indexOf" | "lastIndexOf"
                                | "split" | "regionMatches"
                                    if args.iter().any(|a| a.name.is_some()) =>
                                {
                                    let mut new_args = args.clone();
                                    for a in &mut new_args {
                                        a.name = None;
                                    }
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::new(ExprKind::Member {
                                            object: object.clone(),
                                            field: field.clone(),
                                            null_safe: false })),
                                        args: new_args,
                                        optional: false });
                                    continue;
                                }
                                // 2-arg `contains(needle, ignoreCase)` — the
                                // 1-arg form was rewritten above; this one
                                // must dodge `[array_methods]`' `contains`.
                                "startsWith" | "endsWith" if args.len() == 1 => {
                                    // Free-call form of the [builtins] entry —
                                    // the member path missed for literal
                                    // receivers (measured: `"a".startsWith("a")`
                                    // was "undefined is not callable").
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident(field)),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "startsWith" | "endsWith" if args.len() == 2 => {
                                    let target = if field == "startsWith" {
                                        "__kt_starts_with_ic"
                                    } else {
                                        "__kt_ends_with_ic"
                                    };
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident(target)),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                            args[1].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "find" | "findLast" if args.len() == 1 => {
                                    let target = if field == "find" {
                                        "__kt_find"
                                    } else {
                                        "__kt_find_last"
                                    };
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident(target)),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "contains" if args.len() == 2 => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__kt_contains_ic")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                            args[1].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "sortedByDescending" if args.len() == 1 => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::new(ExprKind::Member {
                                            object: object.clone(),
                                            field: "sort".to_string(),
                                            null_safe: false })),
                                        args: vec![Argument::positional(
                                            kotlin_compare_by_lambda(
                                                args[0].value.clone(),
                                                true,
                                            ),
                                        )],
                                        optional: false });
                                    continue;
                                }
                                // `slice(1..3)` — the range's bounds become the
                                // shared slice's `[from, to)` pair.
                                "slice" | "sliceArray"
                                    if args.len() == 1
                                        && !matches!(
                                            args[0].value.kind,
                                            ExprKind::Range { .. }
                                        ) =>
                                {
                                    // The range arrives as a VALUE
                                    // (`IntRange(0, 2)`, a stored range) —
                                    // materialized array; slice by its bounds.
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__kt_slice_r")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "slice" | "sliceArray" if args.len() == 1
                                    && matches!(args[0].value.kind, ExprKind::Range { .. }) =>
                                {
                                    if let ExprKind::Range { start, end, inclusive } =
                                        &args[0].value.kind
                                    {
                                        let to = if *inclusive {
                                            Expression::new(ExprKind::Binary {
                                                op: BinOp::Add,
                                                left: end.clone(),
                                                right: Box::new(Expression::int(1)) })
                                        } else {
                                            (**end).clone()
                                        };
                                        current = Expression::new(ExprKind::Call {
                                            callee: Box::new(Expression::ident("__kt_slice")),
                                            args: vec![
                                                Argument::positional(*object.clone()),
                                                Argument::positional((**start).clone()),
                                                Argument::positional(to),
                                            ],
                                            optional: false });
                                        continue;
                                    }
                                    continue;
                                }
                                "contentToString" if args.is_empty() => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__kt_tostring")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false });
                                    continue;
                                }
                                "sum" if args.is_empty() => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_sum")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false });
                                    continue;
                                }
                                "reduce" if args.len() == 1 => {
                                    // Kotlin's `reduce` THROWS on empty; the
                                    // shared array HOF answers null.
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__kt_reduce")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "reduce" if args.len() == 1 => {
                                    // Kotlin's `reduce` THROWS on empty; the
                                    // shared array HOF answers null.
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__kt_reduce")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            args[0].clone(),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "fold" if args.len() == 2 => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::new(ExprKind::Member {
                                            object: object.clone(),
                                            field: "fold".to_string(),
                                            null_safe: false })),
                                        args: vec![
                                            Argument::positional(args[1].value.clone()),
                                            Argument::positional(args[0].value.clone()),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "take" if args.len() == 1 => {
                                    let source = kotlin_sequence_source(object)
                                        .unwrap_or_else(|| *object.clone());
                                    if let Some(materialized) =
                                        kotlin_generate_sequence_take(&source, &args[0].value)
                                    {
                                        current = materialized;
                                        continue;
                                    }
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_slice")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            Argument::positional(Expression::int(0)),
                                            Argument::positional(args[0].value.clone()),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "takeWhile" if args.len() == 1 => {
                                    let source = kotlin_sequence_source(object)
                                        .unwrap_or_else(|| *object.clone());
                                    if let Some(materialized) =
                                        kotlin_generate_sequence_take_while(&source, &args[0].value)
                                    {
                                        current = materialized;
                                        continue;
                                    }
                                }
                                "asSequence" if args.is_empty() => {
                                    current = *object.clone();
                                    continue;
                                }
                                "toList" | "toMutableList" if args.is_empty() => {
                                    let source = kotlin_sequence_source(object)
                                        .unwrap_or_else(|| *object.clone());
                                    // Not a materializable sequence: the
                                    // runtime probe handles the rest —
                                    // array→clone, set→values, MAP→entry
                                    // pairs (`Map.toList()` is `List<Pair>`;
                                    // the identity fallback answered the map
                                    // itself).
                                    current = kotlin_materialize_generate_sequence(&source, None)
                                        .unwrap_or_else(|| {
                                            Expression::new(ExprKind::Call {
                                                callee: Box::new(Expression::ident(
                                                    "__kt_map_entry_list",
                                                )),
                                                args: vec![Argument::positional(*object.clone())],
                                                optional: false })
                                        });
                                    continue;
                                }
                                "windowed" if args.len() == 1 => {
                                    let source = kotlin_sequence_source(object)
                                        .unwrap_or_else(|| *object.clone());
                                    if let ExprKind::Lit(Literal::Int(size)) = args[0].value.kind {
                                        if let Some(items) = kotlin_static_array_items(&source) {
                                            if let Some(windowed) =
                                                kotlin_window_static(&items, size as usize)
                                            {
                                                current = windowed;
                                                continue;
                                            }
                                        }
                                    }
                                }
                                "chunked" if args.len() == 1 || args.len() == 2 => {
                                    let source = kotlin_sequence_source(object)
                                        .unwrap_or_else(|| *object.clone());
                                    if let ExprKind::Lit(Literal::Int(size)) = args[0].value.kind {
                                        if let Some(items) = kotlin_static_array_items(&source) {
                                            if args.len() == 2
                                                && kotlin_lambda_is_collection_sum(&args[1].value)
                                            {
                                                let mut sums = Vec::new();
                                                for chunk in items.chunks(size as usize) {
                                                    if let Some(sum) =
                                                        kotlin_sum_static_items(chunk)
                                                    {
                                                        sums.push(sum);
                                                    } else {
                                                        sums.clear();
                                                        break;
                                                    }
                                                }
                                                if !sums.is_empty() || items.is_empty() {
                                                    current = kotlin_array_expr(sums);
                                                    continue;
                                                }
                                            }
                                            if let Some(chunked) =
                                                kotlin_chunk_static(&items, size as usize)
                                            {
                                                current = chunked;
                                                continue;
                                            }
                                        }
                                    }
                                }
                                "zipWithNext" if args.is_empty() => {
                                    let source = kotlin_sequence_source(object)
                                        .unwrap_or_else(|| *object.clone());
                                    if let Some(items) = kotlin_static_array_items(&source) {
                                        current = kotlin_zip_with_next_static(&items);
                                        continue;
                                    }
                                }
                                "drop" if args.len() == 1 => {
                                    let len_expr = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_length")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false });
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_slice")),
                                        args: vec![
                                            Argument::positional(*object.clone()),
                                            Argument::positional(args[0].value.clone()),
                                            Argument::positional(len_expr),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "first" | "firstOrNull" if args.is_empty() => {
                                    // `first()` throws NoSuchElementException on
                                    // empty; only `firstOrNull()` answers null.
                                    let target = if field == "first" {
                                        "__kt_first"
                                    } else {
                                        "__kt_first_or_null"
                                    };
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident(target)),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false });
                                    continue;
                                }
                                "max" | "maxOrNull" if args.is_empty() => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_max")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false });
                                    continue;
                                }
                                "min" | "minOrNull" if args.is_empty() => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_min")),
                                        args: vec![Argument::positional(*object.clone())],
                                        optional: false });
                                    continue;
                                }
                                "filter" | "filterNot" | "map" | "forEach" if args.len() == 1 => {
                                    // Receiver-dispatching adapters: Kotlin's
                                    // `filter` returns a Map ON a Map and a List
                                    // on a List/Set, which no compile-time HOF
                                    // over one shape can do. The old rewrite here
                                    // flattened a Map to raw `[k, v]` pairs (so
                                    // `it.value` read `undefined`) and mapped
                                    // `filterNot` onto `filter`, LOSING the
                                    // negation.
                                    let target = match field.as_str() {
                                        "filter" => "__kt_filter",
                                        "filterNot" => "__kt_filter_not",
                                        "map" => "__kt_map_hof",
                                        _ => "__kt_for_each" };
                                    let mut new_args =
                                        vec![Argument::positional(*object.clone())];
                                    new_args.extend(args);
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident(target)),
                                        args: new_args,
                                        optional: false });
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
                            _ => false };

                        if matches!(
                            &current.kind,
                            ExprKind::Ident(name) if matches!(
                                name.as_str(),
                                "ByteArray" | "CharArray" | "String" | "IntArray" | "LongArray"
                                    | "DoubleArray" | "FloatArray" | "BooleanArray" | "List"
                                    | "MutableList" | "Array" | "IntRange" | "LongRange"
                                    | "CharRange"
                            )
                        ) {
                            current = Expression::new(ExprKind::Call {
                                callee: Box::new(current),
                                args,
                                optional: false });
                            continue;
                        }

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
                                args });
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
                                    // Same receiver-dispatching adapters as the
                                    // sibling arm above — see the comment there.
                                    let target = match field.as_str() {
                                        "filter" => "__kt_filter",
                                        "filterNot" => "__kt_filter_not",
                                        "map" => "__kt_map_hof",
                                        _ => "__kt_for_each" };
                                    let mut new_args =
                                        vec![Argument::positional(*object.clone())];
                                    new_args.extend(args);
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident(target)),
                                        args: new_args,
                                        optional: false });
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
                                optional: false });
                        }
                    }
                    Rule::member_suffix => {
                        let field_id = suffix_inner
                            .into_inner()
                            .find(|p| p.as_rule() == Rule::identifier)
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
                                    null_safe: false });
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
                                    null_safe: false });
                            } else {
                                current = Expression::new(ExprKind::SuperCall {
                                    method: Some(field_id),
                                    args: vec![] });
                            }
                        } else if !next_is_call && field_id == "code" {
                            current = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::ident("__kt_char_code")),
                                args: vec![Argument::positional(current)],
                                optional: false });
                        } else if !next_is_call
                            && is_extension_property(&field_id)
                            && !is_user_property_name(&field_id)
                        {
                            current = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::ident(&field_id)),
                                args: vec![Argument::positional(current)],
                                optional: false });
                        } else if !next_is_call && is_user_property_name(&field_id) {
                            // A property a class in this source declares is an
                            // ordinary member read. The rewrites below match on
                            // SPELLING, so `data class Counter(val values:
                            // MutableList<Int>)` had `a.values` answer the
                            // OBJECT's members via `__dict_values`.
                            current = Expression::new(ExprKind::Member {
                                object: Box::new(current),
                                field: field_id.clone(),
                                null_safe: false });
                        } else {
                            let tuple_property = !next_is_call
                                && kotlin_is_tuple_property_receiver(&current)
                                && matches!(field_id.as_str(), "first" | "second" | "third");
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
                                "first" if tuple_property => {
                                    current = Expression::new(ExprKind::Index {
                                        object: Box::new(current),
                                        index: Box::new(Expression::int(0)),
                                        null_safe: false });
                                }
                                "second" if tuple_property => {
                                    current = Expression::new(ExprKind::Index {
                                        object: Box::new(current),
                                        index: Box::new(Expression::int(1)),
                                        null_safe: false });
                                }
                                "third" if tuple_property => {
                                    current = Expression::new(ExprKind::Index {
                                        object: Box::new(current),
                                        index: Box::new(Expression::int(2)),
                                        null_safe: false });
                                }
                                // Statically UNKNOWN receiver (lambda param,
                                // function result): a Pair/Triple is a tagged
                                // array with no by-name props, so `it.first`
                                // as a bare property read answered undefined.
                                // Runtime probe: array → positional read,
                                // object (a data class declaring `first`) →
                                // property read. Also covers `range.first`.
                                "first" | "second" | "third" if !next_is_call => {
                                    let idx = match field_id.as_str() {
                                        "first" => 0,
                                        "second" => 1,
                                        _ => 2 };
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__kt_tuple_prop")),
                                        args: vec![
                                            Argument::positional(current),
                                            Argument::positional(Expression::new(
                                                ExprKind::Lit(Literal::Str(
                                                    field_id.clone().into(),
                                                )),
                                            )),
                                            Argument::positional(Expression::int(idx)),
                                        ],
                                        optional: false });
                                }
                                "keys" if !next_is_call => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__dict_keys")),
                                        args: vec![Argument::positional(current)],
                                        optional: false });
                                }
                                "values" if !next_is_call => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__dict_values")),
                                        args: vec![Argument::positional(current)],
                                        optional: false });
                                }
                                // No `!next_is_call` guard: `map.entries.find { }`
                                // chains straight into a call, and the skipped
                                // rewrite left a plain `entries` property read
                                // (undefined on a dict).
                                "entries" => {
                                    // Entry OBJECTS — `[k, v]` with `key`/`value`
                                    // properties stamped on — so `e.key` and
                                    // `(k, v)` destructuring read the same thing.
                                    // `__dict_items` yields bare pairs, on which
                                    // `.key` was `undefined` (rendered `NaN`).
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__kt_map_entries")),
                                        args: vec![Argument::positional(current)],
                                        optional: false });
                                }
                                "size" | "length" if !next_is_call => {
                                    current = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_length")),
                                        args: vec![Argument::positional(current)],
                                        optional: false });
                                }
                                "lastIndex" if !next_is_call => {
                                    let len_expr = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__coll_length")),
                                        args: vec![Argument::positional(current)],
                                        optional: false });
                                    current = Expression::new(ExprKind::Binary {
                                        op: BinOp::Sub,
                                        left: Box::new(len_expr),
                                        right: Box::new(Expression::int(1)) });
                                }
                                "indices" if !next_is_call => {
                                    let len_expr = Expression::new(ExprKind::Member {
                                        object: Box::new(current.clone()),
                                        field: "length".to_string(),
                                        null_safe: false });
                                    current = Expression::new(ExprKind::Range {
                                        start: Box::new(Expression::int(0)),
                                        end: Box::new(Expression::new(ExprKind::Binary {
                                            op: BinOp::Sub,
                                            left: Box::new(len_expr),
                                            right: Box::new(Expression::int(1)) })),
                                        inclusive: true });
                                }
                                _ => {
                                    current = Expression::new(ExprKind::Member {
                                        object: Box::new(current),
                                        field: field_id,
                                        null_safe: false });
                                }
                            }
                        }
                    }
                    Rule::safe_call_suffix => {
                        let field_id = suffix_inner
                            .into_inner()
                            .find(|p| p.as_rule() == Rule::identifier)
                            .unwrap()
                            .as_str()
                            .to_string();
                        // `?.size` and friends get the same collection-property
                        // rewrites as the plain member path — a bare property
                        // read on a list/dict answers nothing.
                        // These are SAFE calls — the collection-property
                        // rewrites must keep the null guard (`x?.length` on
                        // a null receiver crashed inside `size_any`).
                        let guarded = |target: &str, recv: Expression| {
                            Expression::new(ExprKind::Ternary {
                                cond: Box::new(Expression::new(ExprKind::Binary {
                                    op: BinOp::Eq,
                                    left: Box::new(recv.clone()),
                                    right: Box::new(Expression::new(ExprKind::Lit(
                                        Literal::Null,
                                    ))) })),
                                then: Box::new(Expression::new(ExprKind::Lit(Literal::Null))),
                                else_: Box::new(Expression::new(ExprKind::Call {
                                    callee: Box::new(Expression::ident(target)),
                                    args: vec![Argument::positional(recv)],
                                    optional: false })) })
                        };
                        match field_id.as_str() {
                            "size" | "length" => {
                                current = guarded("__coll_length", current);
                            }
                            "keys" => {
                                current = guarded("__dict_keys", current);
                            }
                            "values" => {
                                current = guarded("__dict_values", current);
                            }
                            "entries" => {
                                current = guarded("__kt_map_entries", current);
                            }
                            // `?.first`/`?.second`/`?.third` — the tuple-prop
                            // probe is null-tolerant, so it covers the safe
                            // path too (Pairs are arrays with no by-name
                            // props; a bare property read answers undefined).
                            "first" | "second" | "third" if !next_is_call => {
                                let idx = match field_id.as_str() {
                                    "first" => 0,
                                    "second" => 1,
                                    _ => 2 };
                                current = Expression::new(ExprKind::Call {
                                    callee: Box::new(Expression::ident("__kt_tuple_prop")),
                                    args: vec![
                                        Argument::positional(current),
                                        Argument::positional(Expression::new(ExprKind::Lit(
                                            Literal::Str(field_id.clone().into()),
                                        ))),
                                        Argument::positional(Expression::int(idx)),
                                    ],
                                    optional: false });
                            }
                            // A CALL follows (`x?.plus(10)`): the member must
                            // stay a callable lookup — a safe-get probe answers
                            // the property VALUE and "null is not callable".
                            _ if next_is_call => {
                                current = Expression::new(ExprKind::Member {
                                    object: Box::new(current),
                                    field: field_id,
                                    null_safe: true });
                            }
                            _ => {
                                // The shared null-safe Member path gates on
                                // `is_object`, which is FALSE for array-backed
                                // values (map entries, pairs, withIndex tuples)
                                // — `?.key` on a live entry answered null. Route
                                // through a runtime probe instead: null receiver
                                // → null, anything else → dynamic property get.
                                current = Expression::new(ExprKind::Call {
                                    callee: Box::new(Expression::ident("__kt_safe_get")),
                                    args: vec![
                                        Argument::positional(current),
                                        Argument::positional(Expression::new(ExprKind::Lit(
                                            Literal::Str(field_id.into()),
                                        ))),
                                    ],
                                    optional: false });
                            }
                        }
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
                            null_safe: false });
                    }
                    Rule::null_assert_suffix => {
                        current = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident("__kt_not_null_assert")),
                            args: vec![Argument::positional(current)],
                            optional: false });
                    }
                    Rule::inc_suffix => {
                        let op_str = suffix_inner.as_str();
                        let op = if op_str == "++" {
                            UnaryOp::PostInc
                        } else {
                            UnaryOp::PostDec
                        };
                        current = Expression::new(ExprKind::Unary {
                            op,
                            expr: Box::new(current) });
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
                Rule::try_expr => walk_try_expr(inner),
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
                // `fun(x: Int): Int = x` / `fun Int.(): Int { … }` — an
                // anonymous function is a lambda; a RECEIVER becomes the
                // leading `__recv` parameter with the body's `this`/bare
                // members bound to it (the call-site rewrite passes the
                // receiver first).
                Rule::anon_fun => {
                    let mut params: Vec<Param> = Vec::new();
                    let mut body: Vec<Statement> = Vec::new();
                    let mut has_receiver = false;
                    for part in inner.into_inner() {
                        match part.as_rule() {
                            Rule::receiver_prefix => has_receiver = true,
                            Rule::parameter_list => params = walk_parameter_list(part),
                            Rule::function_body_expr => {
                                if let Some(e) = part
                                    .into_inner()
                                    .find(|p| matches!(p.as_rule(), Rule::expr))
                                {
                                    body = vec![Statement::new(StmtKind::Return(Some(
                                        walk_expr(e),
                                    )))];
                                }
                            }
                            Rule::block => body = walk_block_statements(part),
                            _ => {}
                        }
                    }
                    if has_receiver {
                        let mut skip = HashSet::new();
                        kotlin_local_binding_names(&body, &mut skip);
                        for p in &params {
                            skip.insert(p.name.clone());
                        }
                        for stmt in &mut body {
                            kotlin_rewrite_receiver_refs_stmt(stmt, "__recv", &skip);
                        }
                        params.insert(0, kotlin_local_capture_params(&["__recv".to_string()])
                            .pop()
                            .unwrap());
                    }
                    Expression::new(ExprKind::Lambda {
                        params,
                        body: LambdaBody::Block(body),
                        is_async: false,
                        captures: Vec::new() })
                }
                Rule::object_expr => {
                    let mut parent = None;
                    let mut interfaces = Vec::new();
                    let mut members = Vec::new();
                    let mut ctor_args: Vec<Argument> = Vec::new();
                    let mut adapter_passthrough: Option<Expression> = None;
                    for osub in inner.into_inner() {
                        match osub.as_rule() {
                            Rule::inheritance_list => {
                                for spec in osub.into_inner() {
                                    if spec.as_rule() == Rule::inheritance_specifier {
                                        let mut parent_name = String::new();
                                        let mut spec_base_args = Vec::new();
                                        // Parentheses mark the SUPERCLASS, the
                                        // same rule `walk_class_decl` uses:
                                        // `object : Base(), I` extends `Base`.
                                        // Taking the first supertype made
                                        // `object : Callback { … }` extend an
                                        // interface it should have implemented.
                                        let calls_constructor = spec.as_str().contains('(');
                                        for sub in spec.into_inner() {
                                            match sub.as_rule() {
                                                Rule::type_ref => {
                                                    parent_name = type_hint_text(sub.as_str());
                                                }
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
                                                _ => {}
                                            }
                                        }
                                        if !parent_name.is_empty() {
                                            if calls_constructor
                                                && matches!(
                                                    parent_name.as_str(),
                                                    "java.io.FilterInputStream"
                                                        | "java.io.FilterWriter"
                                                )
                                            {
                                                adapter_passthrough =
                                                    Some(Expression::new(ExprKind::New {
                                                        class: Box::new(Expression::ident(
                                                            &parent_name,
                                                        )),
                                                        args: spec_base_args
                                                            .into_iter()
                                                            .map(Argument::positional)
                                                            .collect() }));
                                                continue;
                                            }
                                            if calls_constructor && parent.is_none() {
                                                parent =
                                                    Some(Box::new(Expression::ident(&parent_name)));
                                                ctor_args = spec_base_args
                                                    .into_iter()
                                                    .map(Argument::positional)
                                                    .collect();
                                            } else {
                                                interfaces.push(parent_name);
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
                                                            null_safe: false },
                                                    )),
                                                    value: Box::new(value) }),
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
                                        visibility: Visibility::Public });
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
                    if let Some(expr) = adapter_passthrough {
                        return expr;
                    }
                    Expression::new(ExprKind::New {
                        class: Box::new(Expression::new(ExprKind::ClassExpr {
                            name: None,
                            parent,
                            interfaces,
                            members })),
                        args: ctor_args })
                }
                Rule::if_expr => {
                    let stmt = walk_if_stmt(inner).unwrap();
                    kotlin_if_stmt_to_ternary(stmt)
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
                                    body_expr = kotlin_block_statements_as_expr(stmts);
                                }
                                Rule::statement => {
                                    if let Some(s) = walk_statement(p) {
                                        body_expr = match s.kind {
                                            StmtKind::Expr(e) => e,
                                            StmtKind::Return(Some(e)) => e,
                                            // An `if/else` ARM BODY is a
                                            // value here (`in 90..100 ->
                                            // if (…) "A" else "A+"`).
                                            kind @ StmtKind::If { .. } => {
                                                kotlin_if_stmt_to_ternary(Statement::new(kind))
                                            }
                                            _ => Expression::null() };
                                    }
                                }
                                _ => {}
                            }
                        }

                        let conditions = if is_else { None } else { Some(cond_exprs) };

                        arms.push(MatchArm {
                            conditions,
                            body: body_expr });
                    }

                    Expression::new(ExprKind::Match {
                        subject: Box::new(if predicate_mode {
                            Expression::bool(true)
                        } else {
                            subject
                        }),
                        arms })
                }
                Rule::expr => walk_expr(inner),
                _ => Expression::null() }
        }
        _ => Expression::null() }
}

fn walk_binary_chain(pair: Pair<Rule>, op: BinOp) -> Expression {
    let mut inner = pair.into_inner();
    let mut current = walk_expr(inner.next().unwrap());
    while let Some(_op_pair) = inner.next() {
        let next_expr = walk_expr(inner.next().unwrap());
        current = Expression::new(ExprKind::Binary {
            op,
            left: Box::new(current),
            right: Box::new(next_expr) });
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
    let trimmed = raw.trim().trim_end_matches('?').trim_end();
    let mut depth = 0usize;
    let mut out = String::new();
    for ch in trimmed.chars() {
        match ch {
            '<' => {
                depth += 1;
                if depth == 1 {
                    continue;
                }
            }
            '>' => {
                depth = depth.saturating_sub(1);
                if depth == 0 {
                    continue;
                }
            }
            '?' if depth == 0 => continue,
            _ => {}
        }
        if depth == 0 {
            out.push(ch);
        }
    }
    out.trim().to_string()
}

fn kotlin_nullable_type_hint(raw: &str) -> (Option<String>, bool) {
    let nullable = type_ref_is_nullable(raw);
    let hint = if nullable {
        None
    } else {
        Some(type_hint_text(raw))
    };
    (hint, nullable)
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
        _ => None }
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
                s => s.to_string() };
            Expression::string(&decoded)
        }
        _ => Expression::null() }
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
            _ => folded.push(part) }
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
                right: Box::new(p) });
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
        optional: false })
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
            Rule::str_text | Rule::raw_str_text | Rule::str_lone_dollar => {
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
                    _ => s.to_string() };
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
                            let walked = walk_expr(epair);
                            // A literal range renders as `a..b` in Kotlin; our
                            // ranges lower to materialized arrays, which would
                            // render `[a, …, b]`.
                            if let ExprKind::Range { start, end, inclusive } = &walked.kind {
                                if let (
                                    ExprKind::Lit(Literal::Int(a)),
                                    ExprKind::Lit(Literal::Int(b)),
                                ) = (&start.kind, &end.kind)
                                {
                                    let hi = if *inclusive { *b } else { *b - 1 };
                                    parts.push(Expression::string(&format!("{a}..{hi}")));
                                    continue;
                                }
                            }
                            parts.push(interpolated_part(walked));
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

/// Trailing-lambda TRANSFORM overloads that decompose into `.map`:
/// `chunked(n) { f }` → `chunked(n).map { f }`, `windowed(n, s) { f }` →
/// `windowed(n, s).map { f }`, `joinToString(sep) { f }` →
/// `map { f }.joinToString(sep)` (lambda-only form gets Kotlin's default
/// `", "` separator). Answers None when the last argument is not a lambda.
fn kotlin_lambda_tail_rewrite(
    object: &Expression,
    field: &str,
    args: &[Argument],
) -> Option<Expression> {
    let last_is_lambda = args
        .last()
        .is_some_and(|a| matches!(a.value.kind, ExprKind::Lambda { .. }));
    if !last_is_lambda {
        return None;
    }
    let lambda = args.last().unwrap().clone();
    let member_call = |obj: Expression, name: &str, call_args: Vec<Argument>| {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(obj),
                field: name.to_string(),
                null_safe: false })),
            args: call_args,
            optional: false })
    };
    match field {
        "chunked" | "windowed" if args.len() >= 2 => {
            let inner = member_call(object.clone(), field, args[..args.len() - 1].to_vec());
            Some(member_call(inner, "map", vec![lambda]))
        }
        "joinToString" => {
            let mapped = member_call(object.clone(), "map", vec![lambda]);
            let mut sep_args = args[..args.len() - 1].to_vec();
            if sep_args.is_empty() {
                sep_args.push(Argument::positional(Expression::new(ExprKind::Lit(
                    Literal::Str(", ".into()),
                ))));
            }
            Some(member_call(mapped, "joinToString", sep_args))
        }
        _ => None }
}

fn kotlin_is_tuple_property_receiver(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Tuple(_) => true,
        ExprKind::Ident(name) => KOTLIN_TUPLE_LOCALS.with(|set| set.borrow().contains(name)),
        _ => false }
}

fn kotlin_map_call_expr(object: Expression, transform: Argument) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(object),
            field: "map".to_string(),
            null_safe: false })),
        args: vec![transform],
        optional: false })
}

/// `mapOf(…)` / a Kotlin map literal.
///
/// Plain data. This used to append a synthesised `toString` PROPERTY to the
/// object, which made the map render itself — and put `toString` into the
/// map's own `__keys`, so the map contained a member the program never put
/// there (`{a=1, b=2, toString=…}` once the renderer stopped hiding it).
/// Rendering is `emitter/tostring.rs`'s job; a map is just its entries.
fn create_map_expr(props: Vec<ObjectProperty>) -> Expression {
    Expression::new(ExprKind::Object(
        props
            .into_iter()
            .map(|prop| match prop {
                ObjectProperty::KeyValue { key, value } => ObjectProperty::KeyValue {
                    key: kotlin_key_expr(key),
                    value },
                other => other })
            .collect(),
    ))
}

/// `setOf(…)` / `mutableSetOf(…)`.
///
/// A Kotlin `Set` is a dict whose values are all `true` — the keys ARE the
/// elements, which is what gives `in` its O(1) answer. It carries
/// [`SET_MARKER`] because a `Set` and a `Map` are the same runtime shape and
/// render differently: `[1, 2, 3]` versus `{a=1}`.
fn create_kotlin_set_expr(elems: Vec<Expression>) -> Expression {
    let mut props = vec![ObjectProperty::KeyValue {
        key: Expression::new(ExprKind::Lit(Literal::Str(SET_MARKER.to_string()))),
        value: Expression::bool(true) }];
    let mut seen_literal_keys = HashSet::new();
    for elem in elems {
        let key = kotlin_key_expr(elem.clone());
        if let ExprKind::Lit(Literal::Str(text)) = &key.kind
            && !seen_literal_keys.insert(text.clone())
        {
            continue;
        }
        props.push(ObjectProperty::KeyValue { key, value: elem });
    }
    Expression::new(ExprKind::Object(props))
}
