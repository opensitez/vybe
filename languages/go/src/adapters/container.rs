use vybe_ast::{Argument, ExprKind, Expression};
use vybe_compiler::primitives::namespaces::{NamespaceNode, Subtree};

pub(crate) fn register_tree(root: &mut Subtree) {
    for (name, emit) in [
        ("container.heap.Init", "go.container.heap.Init"),
        ("container.heap.Pop", "go.container.heap.Pop"),
        ("container.heap.Remove", "go.container.heap.Remove"),
        ("container.heap.Fix", "go.container.heap.Fix"),
        (
            "container.heap.remove_prepare",
            "go.container.heap.remove_prepare",
        ),
        ("container.list.New", "go.container.list.New"),
        ("container.ring.New", "go.container.ring.New"),
    ] {
        insert_path(root, name, NamespaceNode::CommonEmit(emit.to_string()));
    }

    let mut list_methods = Subtree::new();
    for method in [
        "Init",
        "Len",
        "Front",
        "Back",
        "PushFront",
        "PushBack",
        "InsertBefore",
        "InsertAfter",
        "Remove",
        "MoveToFront",
        "MoveToBack",
        "MoveBefore",
        "MoveAfter",
        "PushBackList",
        "PushFrontList",
    ] {
        list_methods.insert(
            method.to_string(),
            NamespaceNode::CommonEmit(format!("go.container.list.List.{method}")),
        );
    }
    insert_path(
        root,
        "container.list.List",
        NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics: Subtree::new(),
            methods: list_methods,
            member_returns: [
                ("Init".to_string(), "*__goList".to_string()),
                ("Front".to_string(), "*__goListElement".to_string()),
                ("Back".to_string(), "*__goListElement".to_string()),
                ("PushFront".to_string(), "*__goListElement".to_string()),
                ("PushBack".to_string(), "*__goListElement".to_string()),
                ("InsertBefore".to_string(), "*__goListElement".to_string()),
                ("InsertAfter".to_string(), "*__goListElement".to_string()),
            ]
            .into_iter()
            .collect(),
        },
    );

    let mut element_methods = Subtree::new();
    for method in ["Next", "Prev"] {
        element_methods.insert(
            method.to_string(),
            NamespaceNode::CommonEmit(format!("go.container.list.Element.{method}")),
        );
    }
    insert_path(
        root,
        "container.list.Element",
        NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics: Subtree::new(),
            methods: element_methods,
            member_returns: [
                ("Next".to_string(), "*__goListElement".to_string()),
                ("Prev".to_string(), "*__goListElement".to_string()),
            ]
            .into_iter()
            .collect(),
        },
    );

    let mut ring_methods = Subtree::new();
    for method in ["Next", "Prev", "Len", "Move", "Do", "Link", "Unlink"] {
        ring_methods.insert(
            method.to_string(),
            NamespaceNode::CommonEmit(format!("go.container.ring.Ring.{method}")),
        );
    }
    insert_path(
        root,
        "container.ring.Ring",
        NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics: Subtree::new(),
            methods: ring_methods,
            member_returns: [
                ("Next".to_string(), "*__goRing".to_string()),
                ("Prev".to_string(), "*__goRing".to_string()),
                ("Move".to_string(), "*__goRing".to_string()),
                ("Link".to_string(), "*__goRing".to_string()),
                ("Unlink".to_string(), "*__goRing".to_string()),
            ]
            .into_iter()
            .collect(),
        },
    );
}

pub(crate) fn type_binding(type_name: &str) -> Option<&'static str> {
    match type_name.trim().trim_start_matches('*') {
        "list.List" | "container/list.List" | "container.list.List" => Some("__goList"),
        "list.Element" | "container/list.Element" | "container.list.Element" => {
            Some("__goListElement")
        }
        "ring.Ring" | "container/ring.Ring" | "container.ring.Ring" => Some("__goRing"),
        _ => None,
    }
}

pub(crate) fn rewrite_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    let mapped = match call_name {
        "heap.Init" => "go.container.heap.Init",
        "heap.Pop" => "go.container.heap.Pop",
        "heap.Remove" => "go.container.heap.Remove",
        "heap.Fix" => "go.container.heap.Fix",
        "list.New" => "go.container.list.New",
        "ring.New" => "go.container.ring.New",
        _ => return None,
    };
    Some(go_builtin_call(
        mapped,
        args.iter().map(|arg| arg.value.clone()).collect(),
    ))
}

pub(crate) fn rewrite_method_call(
    receiver_type: &str,
    object: &Expression,
    field: &str,
    args: &[Argument],
) -> Option<Expression> {
    let base = match receiver_type.trim().trim_start_matches('*') {
        "__goList"
            if matches!(
                field,
                "Init"
                    | "Len"
                    | "Front"
                    | "Back"
                    | "PushFront"
                    | "PushBack"
                    | "InsertBefore"
                    | "InsertAfter"
                    | "Remove"
                    | "MoveToFront"
                    | "MoveToBack"
                    | "MoveBefore"
                    | "MoveAfter"
                    | "PushBackList"
                    | "PushFrontList"
            ) =>
        {
            "go.container.list.List"
        }
        "__goListElement" if matches!(field, "Next" | "Prev") => "go.container.list.Element",
        "__goRing"
            if matches!(
                field,
                "Next" | "Prev" | "Len" | "Move" | "Do" | "Link" | "Unlink"
            ) =>
        {
            "go.container.ring.Ring"
        }
        _ => return None,
    };

    let mut values = Vec::with_capacity(args.len() + 1);
    values.push(object.clone());
    values.extend(args.iter().enumerate().map(|(idx, arg)| {
        if base == "go.container.ring.Ring" && field == "Do" && idx == 0 {
            unwrap_callable_cast(arg.value.clone())
        } else {
            arg.value.clone()
        }
    }));
    Some(go_builtin_call(&format!("{base}.{field}"), values))
}

pub(crate) fn rewrite_value_get(
    receiver_type: &str,
    object: Expression,
    field: &str,
) -> Option<Expression> {
    if field != "Value" {
        return None;
    }
    let base = match receiver_type.trim().trim_start_matches('*') {
        "__goRing" => "go.container.ring.Ring.Value.Get",
        _ => return None,
    };
    Some(go_builtin_call(base, vec![object]))
}

pub(crate) fn rewrite_value_set(
    receiver_type: &str,
    object: Expression,
    value: Expression,
    field: &str,
) -> Option<Expression> {
    if field != "Value" {
        return None;
    }
    let base = match receiver_type.trim().trim_start_matches('*') {
        "__goRing" => "go.container.ring.Ring.Value.Set",
        _ => return None,
    };
    Some(go_builtin_call(base, vec![object, value]))
}

pub(crate) fn call_type_hint(name: &str) -> Option<&'static str> {
    match name {
        "go.container.list.New" | "go.container.list.List.Init" => Some("*__goList"),
        "go.container.list.List.Front"
        | "go.container.list.List.Back"
        | "go.container.list.List.PushFront"
        | "go.container.list.List.PushBack"
        | "go.container.list.List.InsertBefore"
        | "go.container.list.List.InsertAfter"
        | "go.container.list.Element.Next"
        | "go.container.list.Element.Prev" => Some("*__goListElement"),
        "go.container.ring.New"
        | "go.container.ring.Ring.Next"
        | "go.container.ring.Ring.Prev"
        | "go.container.ring.Ring.Move"
        | "go.container.ring.Ring.Link"
        | "go.container.ring.Ring.Unlink" => Some("*__goRing"),
        "go.container.list.List.Len" | "go.container.ring.Ring.Len" => Some("int"),
        _ => None,
    }
}

fn go_builtin_call(name: &str, args: Vec<Expression>) -> Expression {
    let name = private_helper_name(name);
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(name)),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

fn unwrap_callable_cast(expr: Expression) -> Expression {
    match expr.kind {
        ExprKind::Cast { expr, .. } => *expr,
        _ => expr,
    }
}

fn private_helper_name(name: &str) -> &str {
    match name {
        "go.container.heap.Init" => "__go_container_heap_init",
        "go.container.heap.Pop" => "__go_container_heap_pop",
        "go.container.heap.Remove" => "__go_container_heap_remove",
        "go.container.heap.Fix" => "__go_container_heap_fix",
        "go.container.heap.remove_prepare" => "__go_container_heap_remove_prepare",
        "go.container.list.New" => "__go_container_list_new",
        "go.container.list.List.Init" => "__go_container_list_init",
        "go.container.list.List.Len" => "__go_container_list_len",
        "go.container.list.List.Front" => "__go_container_list_front",
        "go.container.list.List.Back" => "__go_container_list_back",
        "go.container.list.List.PushFront" => "__go_container_list_push_front",
        "go.container.list.List.PushBack" => "__go_container_list_push_back",
        "go.container.list.List.InsertBefore" => "__go_container_list_insert_before",
        "go.container.list.List.InsertAfter" => "__go_container_list_insert_after",
        "go.container.list.List.Remove" => "__go_container_list_remove",
        "go.container.list.List.MoveToFront" => "__go_container_list_move_to_front",
        "go.container.list.List.MoveToBack" => "__go_container_list_move_to_back",
        "go.container.list.List.MoveBefore" => "__go_container_list_move_before",
        "go.container.list.List.MoveAfter" => "__go_container_list_move_after",
        "go.container.list.List.PushBackList" => "__go_container_list_push_back_list",
        "go.container.list.List.PushFrontList" => "__go_container_list_push_front_list",
        "go.container.list.Element.Next" => "__go_container_list_element_next",
        "go.container.list.Element.Prev" => "__go_container_list_element_prev",
        "go.container.ring.New" => "__go_container_ring_new",
        "go.container.ring.Ring.Next" => "__go_container_ring_next",
        "go.container.ring.Ring.Prev" => "__go_container_ring_prev",
        "go.container.ring.Ring.Len" => "__go_container_ring_len",
        "go.container.ring.Ring.Move" => "__go_container_ring_move",
        "go.container.ring.Ring.Do" => "__go_container_ring_do",
        "go.container.ring.Ring.Link" => "__go_container_ring_link",
        "go.container.ring.Ring.Unlink" => "__go_container_ring_unlink",
        "go.container.ring.Ring.Value.Get" => "__go_container_ring_get_value",
        "go.container.ring.Ring.Value.Set" => "__go_container_ring_set_value",
        _ => name,
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
