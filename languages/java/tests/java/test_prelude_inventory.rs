use std::collections::{BTreeMap, BTreeSet};
use vybe_ast::{
    ArrayElement, CaseCondition, ClassMember, ExprKind, Expression, ObjectProperty, Statement,
    StmtKind,
};

#[test]
#[ignore]
fn java_prelude_inventory_report() {
    let prelude = vybe_language_java::emitter::format_runtime::prelude();
    let names = top_level_names(&prelude);
    let mut families: BTreeMap<&'static str, Vec<String>> = BTreeMap::new();
    for name in &names {
        families
            .entry(prelude_family(name))
            .or_default()
            .push(name.clone());
    }

    println!("java prelude statements: {}", prelude.len());
    println!("java prelude top-level names: {}", names.len());
    for (family, names) in &families {
        println!("{family:>16}: {:>4} {}", names.len(), names.join(", "));
    }

    let samples = [
        (
            "plain println",
            "public class Main { public static void main(String[] args) { System.out.println(\"x\"); } }",
        ),
        (
            "classes only",
            "public class Main { static class Base { int value() { return 1; } } static class Child extends Base { int value() { return 2; } } public static void main(String[] args) { System.out.println(new Child().value()); } }",
        ),
        (
            "string ops",
            "public class Main { public static void main(String[] args) { String s = \"abc\"; System.out.println(s.substring(1).toUpperCase()); } }",
        ),
        (
            "formatter",
            "public class Main { public static void main(String[] args) { System.out.printf(\"%04d%n\", 7); } }",
        ),
        (
            "regex",
            "import java.util.regex.*; public class Main { public static void main(String[] args) { Pattern p = Pattern.compile(\"a+\"); System.out.println(p.matcher(\"aaa\").matches()); } }",
        ),
        (
            "scanner",
            "import java.util.*; public class Main { public static void main(String[] args) { Scanner sc = new Scanner(\"a b\"); System.out.println(sc.next()); } }",
        ),
        (
            "collections",
            "import java.util.*; public class Main { public static void main(String[] args) { ArrayList<Integer> xs = new ArrayList<>(); xs.add(1); System.out.println(xs.size()); } }",
        ),
        (
            "threads",
            "public class Main { public static void main(String[] args) throws Exception { Thread t = new Thread(() -> System.out.println(\"x\")); t.start(); t.join(); } }",
        ),
        (
            "big numbers",
            "import java.math.*; public class Main { public static void main(String[] args) { System.out.println(new BigInteger(\"2\").pow(3)); System.out.println(new BigDecimal(\"1.25\").add(new BigDecimal(\"2.75\"))); } }",
        ),
        (
            "url uri",
            "import java.net.*; public class Main { public static void main(String[] args) throws Exception { URL u = new URL(\"https://example.com/a\"); URI r = new URI(\"x:y\"); System.out.println(u.getHost() + r.isOpaque()); } }",
        ),
    ];

    println!();
    println!("user-code __j/common references after prelude prefix:");
    for (label, source) in samples {
        match sample_usage(source, &names) {
            Ok(usage) => {
                let unresolved: Vec<_> = usage
                    .iter()
                    .filter(|name| name.starts_with("__j") && !names.contains(*name))
                    .cloned()
                    .collect();
                println!(
                    "{label:>16}: {:>3} refs {}",
                    usage.len(),
                    usage.iter().cloned().collect::<Vec<_>>().join(", ")
                );
                if !unresolved.is_empty() {
                    println!("                  unresolved: {}", unresolved.join(", "));
                }
            }
            Err(err) => println!("{label:>16}: parse failed: {err}"),
        }
    }
}

fn sample_usage(
    source: &str,
    prelude_names: &BTreeSet<String>,
) -> Result<BTreeSet<String>, String> {
    let module = vybe_language_java::parse(source)?;
    let prelude_prefix = module
        .body
        .iter()
        .take_while(|stmt| is_prelude_stmt(stmt, prelude_names))
        .count();
    let mut refs = BTreeSet::new();
    for stmt in module.body.iter().skip(prelude_prefix) {
        collect_stmt_refs(stmt, &mut refs);
    }
    Ok(refs
        .into_iter()
        .filter(|name| {
            name.starts_with("__j")
                || name.starts_with("__java")
                || name.starts_with("common:")
                || matches!(name.as_str(), "println" | "print")
        })
        .collect())
}

fn is_prelude_stmt(stmt: &Statement, names: &BTreeSet<String>) -> bool {
    match &stmt.kind {
        StmtKind::FunctionDecl { name, .. } | StmtKind::ClassDecl { name, .. } => {
            names.contains(name)
        }
        StmtKind::VarDecl { declarations, .. } => declarations.iter().any(|decl| {
            matches!(&decl.pattern, vybe_ast::BindingPattern::Ident(name) if names.contains(name))
        }),
        _ => false }
}

fn top_level_names(stmts: &[Statement]) -> BTreeSet<String> {
    let mut out = BTreeSet::new();
    for stmt in stmts {
        match &stmt.kind {
            StmtKind::FunctionDecl { name, .. } | StmtKind::ClassDecl { name, .. } => {
                out.insert(name.clone());
            }
            StmtKind::VarDecl { declarations, .. } => {
                for decl in declarations {
                    if let vybe_ast::BindingPattern::Ident(name) = &decl.pattern {
                        out.insert(name.clone());
                    }
                }
            }
            _ => {}
        }
    }
    out
}

fn prelude_family(name: &str) -> &'static str {
    match name {
        "__j_out" | "__j_buf" => "io-core",
        _ if name.contains("_bd_") => "bigdecimal",
        _ if name.contains("_bi_") => "biginteger",
        _ if name.contains("_regex") || name.contains("_pat_") || name.contains("_m_") => "regex",
        _ if name.contains("_scanner") || name.contains("_scan") => "scanner",
        _ if name.contains("_thread")
            || name.contains("_monitor")
            || name.contains("_process")
            || name.contains("_runtime") =>
        {
            "thread-process"
        }
        _ if name.contains("_list")
            || name.contains("_map")
            || name.contains("_set")
            || name.contains("_queue")
            || name.contains("_deque")
            || name.contains("_collection")
            || name.contains("_array") =>
        {
            "collections"
        }
        _ if name.contains("_url") || name.contains("_uri") => "url-uri",
        _ if name.contains("_b64") || name.contains("_base64") => "base64",
        _ if name.contains("_optional") => "optional",
        _ if name.contains("_stream") || name.contains("_spliterator") => "streams",
        _ if name.contains("_string") || name.contains("_sb_") || name.contains("_sj_") => {
            "strings"
        }
        _ if name.contains("_fmt") || name.contains("_printf") || name == "__j_sprintf" => "format",
        _ if name.contains("_object") || name.contains("_objects") => "objects",
        _ if name.contains("_prop") || name.contains("_system") => "system",
        _ => "misc",
    }
}

fn collect_stmt_refs(stmt: &Statement, out: &mut BTreeSet<String>) {
    match &stmt.kind {
        StmtKind::FunctionDecl { body, .. } | StmtKind::Block(body) => {
            for stmt in body {
                collect_stmt_refs(stmt, out);
            }
        }
        StmtKind::ClassDecl { members, .. } | StmtKind::StructDecl { members, .. } => {
            for member in members {
                collect_member_refs(member, out);
            }
        }
        StmtKind::EnumDecl { body_members, .. } => {
            for member in body_members {
                collect_member_refs(member, out);
            }
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &decl.init {
                    collect_expr_refs(init, out);
                }
            }
        }
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => collect_expr_refs(expr, out),
        StmtKind::Assign { targets, value, .. } => {
            for target in targets {
                collect_expr_refs(target, out);
            }
            collect_expr_refs(value, out);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            collect_expr_refs(target, out);
            collect_expr_refs(value, out);
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            collect_expr_refs(cond, out);
            for stmt in then_body {
                collect_stmt_refs(stmt, out);
            }
            for (cond, body) in elifs {
                collect_expr_refs(cond, out);
                for stmt in body {
                    collect_stmt_refs(stmt, out);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    collect_stmt_refs(stmt, out);
                }
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            collect_expr_refs(cond, out);
            for stmt in body {
                collect_stmt_refs(stmt, out);
            }
            if let Some(body) = else_body {
                for stmt in body {
                    collect_stmt_refs(stmt, out);
                }
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            for stmt in body {
                collect_stmt_refs(stmt, out);
            }
            collect_expr_refs(cond, out);
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                collect_stmt_refs(init, out);
            }
            if let Some(cond) = cond {
                collect_expr_refs(cond, out);
            }
            if let Some(update) = update {
                collect_expr_refs(update, out);
            }
            for stmt in body {
                collect_stmt_refs(stmt, out);
            }
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            collect_expr_refs(iter, out);
            for stmt in body {
                collect_stmt_refs(stmt, out);
            }
            if let Some(body) = else_body {
                for stmt in body {
                    collect_stmt_refs(stmt, out);
                }
            }
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            collect_expr_refs(expr, out);
            for case in cases {
                for condition in &case.conditions {
                    collect_case_condition_refs(condition, out);
                }
                for stmt in &case.body {
                    collect_stmt_refs(stmt, out);
                }
            }
            if let Some(default) = default {
                for stmt in default {
                    collect_stmt_refs(stmt, out);
                }
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
                collect_stmt_refs(stmt, out);
            }
            for catch in catches {
                for stmt in &catch.body {
                    collect_stmt_refs(stmt, out);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    collect_stmt_refs(stmt, out);
                }
            }
            if let Some(body) = finally {
                for stmt in body {
                    collect_stmt_refs(stmt, out);
                }
            }
        }
        StmtKind::Throw {
            expr: Some(expr),
            cause,
        } => {
            collect_expr_refs(expr, out);
            if let Some(cause) = cause {
                collect_expr_refs(cause, out);
            }
        }
        _ => {}
    }
}

fn collect_member_refs(member: &ClassMember, out: &mut BTreeSet<String>) {
    match member {
        ClassMember::Field {
            init: Some(init), ..
        } => collect_expr_refs(init, out),
        ClassMember::Constructor {
            body, base_args, ..
        } => {
            if let Some(base_args) = base_args {
                for arg in base_args {
                    collect_expr_refs(arg, out);
                }
            }
            for stmt in body {
                collect_stmt_refs(stmt, out);
            }
        }
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => collect_stmt_refs(stmt, out),
        _ => {}
    }
}

fn collect_case_condition_refs(condition: &CaseCondition, out: &mut BTreeSet<String>) {
    match condition {
        CaseCondition::Value(expr) | CaseCondition::Comparison { expr, .. } => {
            collect_expr_refs(expr, out)
        }
        CaseCondition::Range { from, to } => {
            collect_expr_refs(from, out);
            collect_expr_refs(to, out);
        }
    }
}

fn collect_array_element_refs(elem: &ArrayElement, out: &mut BTreeSet<String>) {
    if let Some(key) = &elem.key {
        collect_expr_refs(key, out);
    }
    collect_expr_refs(&elem.value, out);
}

fn collect_expr_refs(expr: &Expression, out: &mut BTreeSet<String>) {
    match &expr.kind {
        ExprKind::Ident(name) => {
            out.insert(name.clone());
        }
        ExprKind::Call { callee, args, .. } => {
            collect_expr_refs(callee, out);
            for arg in args {
                collect_expr_refs(&arg.value, out);
            }
        }
        ExprKind::New { class, args } => {
            collect_expr_refs(class, out);
            for arg in args {
                collect_expr_refs(&arg.value, out);
            }
        }
        ExprKind::Member { object, .. } => collect_expr_refs(object, out),
        ExprKind::Index { object, index, .. } => {
            collect_expr_refs(object, out);
            collect_expr_refs(index, out);
        }
        ExprKind::Binary { left, right, .. } => {
            collect_expr_refs(left, out);
            collect_expr_refs(right, out);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::Spread(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr)
        | ExprKind::RefLoad(expr) => collect_expr_refs(expr, out),
        ExprKind::Yield(Some(expr)) => collect_expr_refs(expr, out),
        ExprKind::Assign { target, value } => {
            collect_expr_refs(target, out);
            collect_expr_refs(value, out);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            collect_expr_refs(cond, out);
            collect_expr_refs(then, out);
            collect_expr_refs(else_, out);
        }
        ExprKind::Array(elems) => {
            for elem in elems {
                collect_array_element_refs(elem, out);
            }
        }
        ExprKind::Tuple(elems) | ExprKind::Set(elems) | ExprKind::Sequence(elems) => {
            for elem in elems {
                collect_expr_refs(elem, out);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value, .. } => {
                        collect_expr_refs(key, out);
                        collect_expr_refs(value, out);
                    }
                    ObjectProperty::Spread(expr) => collect_expr_refs(expr, out),
                    ObjectProperty::Method { value, .. }
                    | ObjectProperty::Accessor { value, .. } => collect_stmt_refs(value, out),
                    ObjectProperty::Computed { key, value } => {
                        collect_expr_refs(key, out);
                        collect_expr_refs(value, out);
                    }
                    ObjectProperty::Shorthand(name) => {
                        out.insert(name.clone());
                    }
                }
            }
        }
        ExprKind::Lambda { body, .. } => match body {
            vybe_ast::LambdaBody::Expr(expr) => collect_expr_refs(expr, out),
            vybe_ast::LambdaBody::Block(stmts) => {
                for stmt in stmts {
                    collect_stmt_refs(stmt, out);
                }
            }
        },
        ExprKind::ClassExpr {
            parent, members, ..
        } => {
            if let Some(parent) = parent {
                collect_expr_refs(parent, out);
            }
            for member in members {
                collect_member_refs(member, out);
            }
        }
        ExprKind::StaticAccess { class, member } => {
            collect_expr_refs(class, out);
            collect_expr_refs(member, out);
        }
        ExprKind::IsType { expr, .. } => collect_expr_refs(expr, out),
        _ => {}
    }
}
