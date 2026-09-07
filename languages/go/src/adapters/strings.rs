use vybe_ast::{Argument, ExprKind, Expression, ObjectProperty};
use vybe_compiler::primitives::namespaces::{NamespaceNode, Subtree};

pub(crate) fn register_tree(root: &mut Subtree) {
    register_type(
        root,
        "__goReplacer",
        &["Replace", "ReplaceCascade", "WriteString"],
    );
    register_type(
        root,
        "strings.Replacer",
        &["Replace", "ReplaceCascade", "WriteString"],
    );

    for (name, emit) in [
        ("strings.ToUpper", "strings.to_upper"),
        ("strings.ToLower", "strings.to_lower"),
        ("strings.TrimSpace", "strings.trim"),
        ("strings.Contains", "str_contains"),
        ("strings.HasPrefix", "str_starts_with"),
        ("strings.HasSuffix", "str_ends_with"),
        ("strings.Index", "strings.index_of"),
        ("strings.Repeat", "str_repeat"),
        ("strings.Compare", "str_compare"),
        ("strings.Split", "strings.split"),
        ("strings.Join", "collections.join"),
        ("strings.LastIndex", "str_last_index_of"),
        ("strings.Title", "str_to_upper"),
        ("strings.TrimPrefix", "go.strings.TrimPrefix"),
        ("strings.TrimSuffix", "go.strings.TrimSuffix"),
        ("strings.CutPrefix", "go.strings.CutPrefix"),
        ("strings.CutSuffix", "go.strings.CutSuffix"),
        ("strings.Cut", "go.strings.Cut"),
        ("strings.Replace", "go.strings.Replace"),
        ("strings.ReplaceAll", "go.strings.ReplaceAll"),
        ("strings.ContainsRune", "go.strings.ContainsRune"),
        ("strings.ContainsAny", "go.strings.ContainsAny"),
        ("strings.ContainsFunc", "go.strings.ContainsFunc"),
        ("strings.IndexByte", "go.strings.IndexByte"),
        ("strings.IndexRune", "go.strings.IndexRune"),
        ("strings.IndexAny", "go.strings.IndexAny"),
        ("strings.IndexFunc", "go.strings.IndexFunc"),
        ("strings.LastIndexByte", "go.strings.LastIndexByte"),
        ("strings.LastIndexAny", "go.strings.LastIndexAny"),
        ("strings.LastIndexFunc", "go.strings.LastIndexFunc"),
        ("strings.TrimLeft", "go.strings.TrimLeft"),
        ("strings.TrimRight", "go.strings.TrimRight"),
        ("strings.Trim", "go.strings.Trim"),
        ("strings.EqualFold", "go.strings.EqualFold"),
        ("strings.Count", "go.strings.Count"),
        ("strings.ToValidUTF8", "go.strings.ToValidUTF8"),
        ("strings.Map", "go.strings.Map"),
        ("strings.Fields", "go.strings.Fields"),
        ("strings.FieldsFunc", "go.strings.FieldsFunc"),
        ("strings.SplitN", "go.strings.SplitN"),
        ("strings.SplitAfter", "go.strings.SplitAfter"),
        ("strings.SplitAfterN", "go.strings.SplitAfterN"),
        ("strings.NewReplacer", "go.strings.NewReplacer"),
    ] {
        insert_path(root, name, NamespaceNode::CommonEmit(emit.to_string()));
    }
}

pub(crate) fn type_binding(type_name: &str) -> Option<&'static str> {
    match type_name.trim().trim_start_matches('*') {
        "strings.Replacer" => Some("__goReplacer"),
        _ => None,
    }
}

pub(crate) fn rewrite_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    let canonical = call_name.strip_prefix("go.").unwrap_or(call_name);
    let mapped = match canonical {
        "strings.TrimPrefix" => "go.strings.TrimPrefix",
        "strings.TrimSuffix" => "go.strings.TrimSuffix",
        "strings.CutPrefix" => "go.strings.CutPrefix",
        "strings.CutSuffix" => "go.strings.CutSuffix",
        "strings.Cut" => "go.strings.Cut",
        "strings.Replace" if args.len() == 4 => "go.strings.Replace",
        "strings.ReplaceAll" if args.len() == 3 => "go.strings.ReplaceAll",
        "strings.ContainsRune" => "go.strings.ContainsRune",
        "strings.ContainsAny" => "go.strings.ContainsAny",
        "strings.ContainsFunc" => "go.strings.ContainsFunc",
        "strings.IndexByte" => "go.strings.IndexByte",
        "strings.IndexRune" => "go.strings.IndexRune",
        "strings.IndexAny" => "go.strings.IndexAny",
        "strings.IndexFunc" => "go.strings.IndexFunc",
        "strings.LastIndexByte" => "go.strings.LastIndexByte",
        "strings.LastIndexAny" => "go.strings.LastIndexAny",
        "strings.LastIndexFunc" => "go.strings.LastIndexFunc",
        "strings.TrimLeft" => "go.strings.TrimLeft",
        "strings.TrimRight" => "go.strings.TrimRight",
        "strings.Trim" => "go.strings.Trim",
        "strings.EqualFold" => "go.strings.EqualFold",
        "strings.Count" => "go.strings.Count",
        "strings.ToValidUTF8" => "go.strings.ToValidUTF8",
        "strings.Map" => "go.strings.Map",
        "strings.Fields" => "go.strings.Fields",
        "strings.FieldsFunc" => "go.strings.FieldsFunc",
        "strings.SplitN" => "go.strings.SplitN",
        "strings.SplitAfter" => "go.strings.SplitAfter",
        "strings.SplitAfterN" => "go.strings.SplitAfterN",
        "strings.NewReplacer" => return Some(replacer_object(args)),
        _ => return None,
    };
    Some(call(private_helper_name(mapped), arg_values(args)))
}

fn private_helper_name(name: &str) -> &str {
    match name {
        "go.strings.TrimPrefix" => "__go_strings_trim_prefix",
        "go.strings.TrimSuffix" => "__go_strings_trim_suffix",
        "go.strings.CutPrefix" => "__go_strings_cut_prefix",
        "go.strings.CutSuffix" => "__go_strings_cut_suffix",
        "go.strings.Cut" => "__go_strings_cut",
        "go.strings.Replace" => "__go_strings_replace",
        "go.strings.ReplaceAll" => "__go_strings_replace_all",
        "go.strings.ContainsRune" => "__go_strings_contains_rune",
        "go.strings.ContainsAny" => "__go_strings_contains_any",
        "go.strings.ContainsFunc" => "__go_strings_contains_func",
        "go.strings.IndexByte" => "__go_strings_index_byte",
        "go.strings.IndexRune" => "__go_strings_index_rune",
        "go.strings.IndexAny" => "__go_strings_index_any",
        "go.strings.IndexFunc" => "__go_strings_index_func",
        "go.strings.LastIndexByte" => "__go_strings_last_index_byte",
        "go.strings.LastIndexAny" => "__go_strings_last_index_any",
        "go.strings.LastIndexFunc" => "__go_strings_last_index_func",
        "go.strings.TrimLeft" => "__go_strings_trim_left",
        "go.strings.TrimRight" => "__go_strings_trim_right",
        "go.strings.Trim" => "__go_strings_trim",
        "go.strings.EqualFold" => "__go_strings_equal_fold",
        "go.strings.Count" => "__go_strings_count",
        "go.strings.ToValidUTF8" => "__go_strings_to_valid_utf8",
        "go.strings.Map" => "__go_strings_map",
        "go.strings.Fields" => "__go_strings_fields",
        "go.strings.FieldsFunc" => "__go_strings_fields_func",
        "go.strings.SplitN" => "__go_strings_split_n",
        "go.strings.SplitAfter" => "__go_strings_split_after",
        "go.strings.SplitAfterN" => "__go_strings_split_after_n",
        _ => name,
    }
}

pub(crate) fn call_type_hint(name: &str) -> Option<&'static str> {
    match name {
        "strings.ToUpper"
        | "strings.ToLower"
        | "strings.TrimSpace"
        | "strings.Repeat"
        | "strings.Join"
        | "strings.Title"
        | "go.strings.TrimPrefix"
        | "go.strings.TrimSuffix"
        | "go.strings.Replace"
        | "go.strings.ReplaceAll"
        | "go.strings.TrimLeft"
        | "go.strings.TrimRight"
        | "go.strings.Trim"
        | "go.strings.ToValidUTF8"
        | "go.strings.Map"
        | "go.__goReplacer.Replace"
        | "go.__goReplacer.ReplaceCascade"
        | "go.strings.Replacer.Replace"
        | "go.strings.Replacer.ReplaceCascade" => Some("string"),
        "strings.Split"
        | "go.strings.Fields"
        | "go.strings.FieldsFunc"
        | "go.strings.SplitN"
        | "go.strings.SplitAfter"
        | "go.strings.SplitAfterN" => Some("[]string"),
        "strings.Contains"
        | "strings.HasPrefix"
        | "strings.HasSuffix"
        | "go.strings.EqualFold"
        | "go.strings.ContainsRune"
        | "go.strings.ContainsAny"
        | "go.strings.ContainsFunc" => Some("bool"),
        "strings.Index"
        | "strings.LastIndex"
        | "strings.Compare"
        | "go.strings.Count"
        | "go.strings.IndexByte"
        | "go.strings.IndexRune"
        | "go.strings.IndexAny"
        | "go.strings.IndexFunc"
        | "go.strings.LastIndexByte"
        | "go.strings.LastIndexAny"
        | "go.strings.LastIndexFunc" => Some("int"),
        _ => None,
    }
}

fn replacer_object(args: &[Argument]) -> Expression {
    pointer_arg(Expression::new(ExprKind::Cast {
        expr: Box::new(Expression::new(ExprKind::Object(vec![
            ObjectProperty::KeyValue {
                key: Expression::string("pairs"),
                value: Expression::new(ExprKind::Array(
                    args.iter()
                        .map(|arg| vybe_ast::ArrayElement {
                            key: None,
                            value: arg.value.clone(),
                            spread: arg.spread,
                            by_ref: false,
                        })
                        .collect(),
                )),
            },
        ]))),
        type_name: "__goReplacer".to_string(),
    }))
}

fn register_type(root: &mut Subtree, name: &str, methods: &[&str]) {
    let mut method_tree = Subtree::new();
    for method in methods {
        method_tree.insert(
            (*method).to_string(),
            NamespaceNode::CommonEmit(format!("go.{name}.{method}")),
        );
    }
    insert_path(
        root,
        name,
        NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics: Subtree::new(),
            methods: method_tree,
            member_returns: Default::default(),
        },
    );
}

fn call(name: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(name)),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

fn pointer_arg(expr: Expression) -> Expression {
    match expr.kind {
        ExprKind::RefOf(_)
        | ExprKind::Unary {
            op: vybe_ast::UnaryOp::AddrOf,
            ..
        } => expr,
        _ => Expression::new(ExprKind::Unary {
            op: vybe_ast::UnaryOp::AddrOf,
            expr: Box::new(expr),
        }),
    }
}

fn arg_values(args: &[Argument]) -> Vec<Expression> {
    args.iter().map(|arg| arg.value.clone()).collect()
}

fn insert_path(root: &mut Subtree, path: &str, node: NamespaceNode) {
    let mut segments: Vec<&str> = path.split('.').collect();
    let Some(leaf) = segments.pop() else {
        return;
    };
    let mut cursor = root;
    for seg in segments {
        let entry = cursor
            .entry(seg.to_string())
            .or_insert_with(|| NamespaceNode::Namespace(Subtree::new()));
        let NamespaceNode::Namespace(children) = entry else {
            return;
        };
        cursor = children;
    }
    cursor.insert(leaf.to_string(), node);
}
