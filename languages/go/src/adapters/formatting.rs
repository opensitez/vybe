use std::collections::HashMap;
use vybe_ast::{BinOp, ExprKind, Expression};
use vybe_compiler::primitives::namespaces::{NamespaceNode, Subtree};

pub(crate) fn register_tree(root: &mut Subtree) {
    for (name, emit) in [
        ("fmt.Println", "go.fmt_println"),
        ("fmt.Print", "go.fmt_print"),
        ("fmt.Printf", "go.fmt_printf"),
        ("fmt.Sprint", "go.fmt_sprint"),
        ("fmt.Sprintf", "go.fmt_sprintf"),
        ("fmt.Errorf", "go.errors_errorf"),
    ] {
        insert_path(root, name, NamespaceNode::CommonEmit(emit.to_string()));
    }
}

#[derive(Clone, Copy)]
pub(crate) enum GoFmtArgRewrite {
    Pointer,
    String,
    Quote,
    TypeName,
    GoValue { field_names: bool },
}

pub(crate) fn rewrite_go_format_literal(fmt: &str) -> (String, HashMap<usize, GoFmtArgRewrite>) {
    let mut out = String::new();
    let mut rewrites = HashMap::new();
    let chars: Vec<char> = fmt.chars().collect();
    let mut i = 0;
    let mut arg_idx = 0usize;

    while i < chars.len() {
        let ch = chars[i];
        if ch != '%' {
            out.push(ch);
            i += 1;
            continue;
        }
        if i + 1 < chars.len() && chars[i + 1] == '%' {
            out.push_str("%%");
            i += 2;
            continue;
        }

        out.push('%');
        i += 1;
        while i < chars.len() {
            let spec = chars[i];
            if spec.is_ascii_alphabetic() {
                match spec {
                    's' => {
                        out.push('s');
                        rewrites.insert(arg_idx, GoFmtArgRewrite::String);
                    }
                    't' | 'v' => {
                        let field_names = spec == 'v' && out.ends_with("%+");
                        if spec == 'v' && out.ends_with("%+") {
                            out.pop();
                        }
                        out.push('s');
                        rewrites.insert(
                            arg_idx,
                            if spec == 'v' {
                                GoFmtArgRewrite::GoValue { field_names }
                            } else {
                                GoFmtArgRewrite::String
                            },
                        );
                    }
                    'q' => {
                        out.push('s');
                        rewrites.insert(arg_idx, GoFmtArgRewrite::Quote);
                    }
                    'T' => {
                        out.push('s');
                        rewrites.insert(arg_idx, GoFmtArgRewrite::TypeName);
                    }
                    'p' => {
                        out.push('s');
                        rewrites.insert(arg_idx, GoFmtArgRewrite::Pointer);
                    }
                    _ => out.push(spec),
                }
                arg_idx += 1;
                i += 1;
                break;
            }
            if spec == '*' {
                arg_idx += 1;
            }
            out.push(spec);
            i += 1;
        }
    }

    (out, rewrites)
}

pub(crate) fn fmt_pointer_expr(value: Expression) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(value),
            right: Box::new(Expression::null()),
        })),
        then: Box::new(Expression::string("0x0")),
        else_: Box::new(Expression::string("0x1")),
    })
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
