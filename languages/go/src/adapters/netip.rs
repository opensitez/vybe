use vybe_ast::{Argument, ExprKind, Expression, ObjectProperty};
use vybe_compiler::primitives::namespaces::{NamespaceNode, Subtree};

pub(crate) fn register_tree(root: &mut Subtree) {
    for (name, methods, returns) in [
        (
            "netip.Addr",
            &[
                "String",
                "Is4",
                "Is6",
                "Is4In6",
                "IsValid",
                "IsUnspecified",
                "IsLoopback",
                "IsPrivate",
                "IsGlobalUnicast",
                "IsLinkLocalUnicast",
                "IsMulticast",
                "Unmap",
                "WithZone",
                "Zone",
                "Compare",
                "Equal",
                "Less",
                "AsSlice",
                "As16",
                "Next",
                "Prev",
            ][..],
            &[
                ("String", "string"),
                ("Is4", "bool"),
                ("Is6", "bool"),
                ("Is4In6", "bool"),
                ("IsValid", "bool"),
                ("IsUnspecified", "bool"),
                ("IsLoopback", "bool"),
                ("IsPrivate", "bool"),
                ("IsGlobalUnicast", "bool"),
                ("IsLinkLocalUnicast", "bool"),
                ("IsMulticast", "bool"),
                ("Unmap", "__goNetipAddr"),
                ("WithZone", "__goNetipAddr"),
                ("Zone", "string"),
                ("Compare", "int"),
                ("Equal", "bool"),
                ("Less", "bool"),
                ("AsSlice", "[]byte"),
                ("As16", "[]byte"),
                ("Next", "__goNetipAddr"),
                ("Prev", "__goNetipAddr"),
            ][..],
        ),
        (
            "netip.Prefix",
            &[
                "String",
                "Bits",
                "IsValid",
                "Addr",
                "Masked",
                "Contains",
                "Overlaps",
                "ContainsPrefix",
            ][..],
            &[
                ("String", "string"),
                ("Bits", "int"),
                ("IsValid", "bool"),
                ("Addr", "__goNetipAddr"),
                ("Masked", "__goNetipPrefix"),
                ("Contains", "bool"),
                ("Overlaps", "bool"),
                ("ContainsPrefix", "bool"),
            ][..],
        ),
        (
            "netip.AddrPort",
            &["String", "Addr", "Port"][..],
            &[
                ("String", "string"),
                ("Addr", "__goNetipAddr"),
                ("Port", "int"),
            ][..],
        ),
    ] {
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
                member_returns: returns
                    .iter()
                    .map(|(member, ty)| ((*member).to_string(), (*ty).to_string()))
                    .collect(),
            },
        );
    }

    for name in [
        "netip.ParseAddr",
        "netip.MustParseAddr",
        "netip.IPv4",
        "netip.AddrFromSlice",
        "netip.ParsePrefix",
        "netip.MustParsePrefix",
        "netip.PrefixFrom",
        "netip.ParseAddrPort",
        "netip.AddrPortFrom",
    ] {
        insert_path(root, name, NamespaceNode::CommonEmit(format!("go.{name}")));
    }
}

pub(crate) fn type_binding(type_name: &str) -> Option<&'static str> {
    match type_name.trim() {
        "netip.Addr" => Some("__goNetipAddr"),
        "netip.Prefix" => Some("__goNetipPrefix"),
        "netip.AddrPort" => Some("__goNetipAddrPort"),
        _ => None,
    }
}

pub(crate) fn zero_value(type_name: &str) -> Option<Expression> {
    match type_name.trim().to_ascii_lowercase().as_str() {
        "__gonetipaddr" => Some(addr_object(
            Expression::string(""),
            Expression::bool(false),
            Expression::bool(false),
            Expression::bool(false),
            Expression::string(""),
            Expression::bool(false),
        )),
        "__gonetipprefix" => Some(prefix_object(
            zero_value("__goNetipAddr").unwrap_or_else(Expression::null),
            Expression::int(0),
            Expression::bool(false),
        )),
        "__gonetipaddrport" => Some(addr_port_object(
            zero_value("__goNetipAddr").unwrap_or_else(Expression::null),
            Expression::int(0),
            Expression::bool(false),
        )),
        _ => None,
    }
}

pub(crate) fn call_type_hint(name: &str) -> Option<&'static str> {
    match name {
        "go.netip.MustParseAddr"
        | "go.netip.IPv4"
        | "go.netip.Addr.Unmap"
        | "go.netip.Addr.WithZone"
        | "go.netip.Addr.Next"
        | "go.netip.Addr.Prev"
        | "go.netip.Prefix.Addr"
        | "go.netip.AddrPort.Addr" => Some("__goNetipAddr"),
        "go.netip.MustParsePrefix" => Some("__goNetipPrefix"),
        "go.netip.Prefix.Masked" => Some("__goNetipPrefix"),
        "go.netip.AddrPortFrom" => Some("__goNetipAddrPort"),
        "go.netip.Addr.String"
        | "go.netip.Addr.Zone"
        | "go.netip.Prefix.String"
        | "go.netip.AddrPort.String" => Some("string"),
        "go.netip.Addr.Compare" | "go.netip.Prefix.Bits" | "go.netip.AddrPort.Port" => Some("int"),
        "go.netip.Addr.Is4"
        | "go.netip.Addr.Is6"
        | "go.netip.Addr.Is4In6"
        | "go.netip.Addr.IsValid"
        | "go.netip.Addr.IsUnspecified"
        | "go.netip.Addr.IsLoopback"
        | "go.netip.Addr.IsPrivate"
        | "go.netip.Addr.IsGlobalUnicast"
        | "go.netip.Addr.IsLinkLocalUnicast"
        | "go.netip.Addr.IsMulticast"
        | "go.netip.Addr.Equal"
        | "go.netip.Addr.Less"
        | "go.netip.Prefix.IsValid"
        | "go.netip.Prefix.Contains"
        | "go.netip.Prefix.Overlaps"
        | "go.netip.Prefix.ContainsPrefix" => Some("bool"),
        "go.netip.Addr.AsSlice" | "go.netip.Addr.As16" => Some("[]byte"),
        "go.netip.ParseAddr"
        | "go.netip.AddrFromSlice"
        | "go.netip.ParsePrefix"
        | "go.netip.PrefixFrom"
        | "go.netip.ParseAddrPort" => Some("tuple"),
        _ => None,
    }
}

pub(crate) fn member_type(type_name: &str, field: &str) -> Option<&'static str> {
    match (
        type_name
            .trim()
            .trim_start_matches('*')
            .trim_start_matches('^')
            .trim(),
        field,
    ) {
        ("__goNetipAddr" | "netip.Addr", "s" | "zone") => Some("string"),
        ("__goNetipAddr" | "netip.Addr", "is4" | "is6" | "is4in6" | "valid") => Some("bool"),
        ("__goNetipPrefix" | "netip.Prefix", "addr") => Some("__goNetipAddr"),
        ("__goNetipPrefix" | "netip.Prefix", "bits") => Some("int"),
        ("__goNetipPrefix" | "netip.Prefix", "valid") => Some("bool"),
        ("__goNetipAddrPort" | "netip.AddrPort", "addr") => Some("__goNetipAddr"),
        ("__goNetipAddrPort" | "netip.AddrPort", "port") => Some("int"),
        ("__goNetipAddrPort" | "netip.AddrPort", "valid") => Some("bool"),
        _ => None,
    }
}

pub(crate) fn tuple_type_hints(expr: &Expression) -> Option<Vec<Option<String>>> {
    match &expr.kind {
        ExprKind::Call { callee, .. } => match go_expr_call_name(callee).as_deref()? {
            "go.netip.ParseAddr" => Some(vec![
                Some("__goNetipAddr".to_string()),
                Some("error".to_string()),
            ]),
            "go.netip.AddrFromSlice" => Some(vec![
                Some("__goNetipAddr".to_string()),
                Some("bool".to_string()),
            ]),
            "go.netip.ParsePrefix" | "go.netip.PrefixFrom" => Some(vec![
                Some("__goNetipPrefix".to_string()),
                Some("error".to_string()),
            ]),
            "go.netip.ParseAddrPort" => Some(vec![
                Some("__goNetipAddrPort".to_string()),
                Some("error".to_string()),
            ]),
            _ => None,
        },
        _ => None,
    }
}

pub(crate) fn rewrite_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    let mapped = match call_name {
        "netip.ParseAddr" => "go.netip.ParseAddr",
        "netip.MustParseAddr" => "go.netip.MustParseAddr",
        "netip.IPv4" => "go.netip.IPv4",
        "netip.AddrFromSlice" => "go.netip.AddrFromSlice",
        "netip.ParsePrefix" => "go.netip.ParsePrefix",
        "netip.MustParsePrefix" => "go.netip.MustParsePrefix",
        "netip.PrefixFrom" => "go.netip.PrefixFrom",
        "netip.ParseAddrPort" => "go.netip.ParseAddrPort",
        "netip.AddrPortFrom" => "go.netip.AddrPortFrom",
        _ => return None,
    };
    Some(call(
        mapped,
        args.iter().map(|arg| arg.value.clone()).collect(),
    ))
}

pub(crate) fn rewrite_method_call(
    receiver: Expression,
    receiver_type: &str,
    method: &str,
    args: &[Argument],
) -> Option<Expression> {
    let ty = receiver_type
        .trim()
        .trim_start_matches('*')
        .trim_start_matches('^')
        .trim();
    let surface = match ty {
        "__goNetipAddr" | "netip.Addr" if has_method(ty, method) => "go.netip.Addr",
        "__goNetipPrefix" | "netip.Prefix" if has_method(ty, method) => "go.netip.Prefix",
        "__goNetipAddrPort" | "netip.AddrPort" if has_method(ty, method) => "go.netip.AddrPort",
        _ => return None,
    };
    let mut values = Vec::with_capacity(args.len() + 1);
    values.push(receiver_obj(receiver, receiver_type));
    values.extend(args.iter().map(|arg| arg.value.clone()));
    Some(call(&format!("{surface}.{method}"), values))
}

pub(crate) fn has_method(type_name: &str, method: &str) -> bool {
    matches!(
        (
            type_name
                .trim()
                .trim_start_matches('*')
                .trim_start_matches('^')
                .trim(),
            method
        ),
        (
            "__goNetipAddr" | "netip.Addr",
            "String"
                | "Is4"
                | "Is6"
                | "Is4In6"
                | "IsValid"
                | "IsUnspecified"
                | "IsLoopback"
                | "IsPrivate"
                | "IsGlobalUnicast"
                | "IsLinkLocalUnicast"
                | "IsMulticast"
                | "Unmap"
                | "WithZone"
                | "Zone"
                | "Compare"
                | "Equal"
                | "Less"
                | "AsSlice"
                | "As16"
                | "Next"
                | "Prev"
        ) | (
            "__goNetipPrefix" | "netip.Prefix",
            "String"
                | "Bits"
                | "IsValid"
                | "Addr"
                | "Masked"
                | "Contains"
                | "Overlaps"
                | "ContainsPrefix"
        ) | (
            "__goNetipAddrPort" | "netip.AddrPort",
            "String" | "Addr" | "Port"
        )
    )
}

fn addr_object(
    s: Expression,
    is4: Expression,
    is6: Expression,
    is4in6: Expression,
    zone: Expression,
    valid: Expression,
) -> Expression {
    typed_object(
        "__goNetipAddr",
        vec![
            ("s", s),
            ("is4", is4),
            ("is6", is6),
            ("is4in6", is4in6),
            ("zone", zone),
            ("valid", valid),
        ],
    )
}

fn prefix_object(addr: Expression, bits: Expression, valid: Expression) -> Expression {
    typed_object(
        "__goNetipPrefix",
        vec![("addr", addr), ("bits", bits), ("valid", valid)],
    )
}

fn addr_port_object(addr: Expression, port: Expression, valid: Expression) -> Expression {
    typed_object(
        "__goNetipAddrPort",
        vec![("addr", addr), ("port", port), ("valid", valid)],
    )
}

fn receiver_obj(receiver: Expression, receiver_type: &str) -> Expression {
    if receiver_type.trim().starts_with('*') {
        Expression::new(ExprKind::RefLoad(Box::new(receiver)))
    } else {
        receiver
    }
}

fn call(name: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(name)),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

fn typed_object(type_name: &str, fields: Vec<(&str, Expression)>) -> Expression {
    Expression::new(ExprKind::Cast {
        expr: Box::new(Expression::new(ExprKind::Object(
            fields
                .into_iter()
                .map(|(name, value)| ObjectProperty::KeyValue {
                    key: Expression::string(name),
                    value,
                })
                .collect(),
        ))),
        type_name: type_name.to_string(),
    })
}

fn go_expr_call_name(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => Some(name.clone()),
        ExprKind::Member { object, field, .. } => {
            let prefix = go_expr_call_name(object)?;
            Some(format!("{prefix}.{field}"))
        }
        ExprKind::Cast { expr, .. } => go_expr_call_name(expr),
        _ => None,
    }
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
