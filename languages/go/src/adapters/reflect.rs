use vybe_compiler::primitives::namespaces::{NamespaceNode, Subtree};

pub(crate) fn register_tree(root: &mut Subtree) {
    for (name, emit) in [
        ("reflect.TypeOf", "go.reflect_typeof"),
        ("reflect.ValueOf", "go.reflect_valueof"),
    ] {
        insert_path(root, name, NamespaceNode::CommonEmit(emit.to_string()));
    }

    register_type(
        root,
        "reflect.Type",
        &["Kind", "Name", "NumMethod", "MethodByName"],
    );
    register_type(
        root,
        "__goReflectType",
        &["Kind", "Name", "NumMethod", "MethodByName"],
    );
    register_type(
        root,
        "reflect.Value",
        &[
            "Kind",
            "Name",
            "Interface",
            "Int",
            "Uint",
            "Float",
            "Bool",
            "String",
            "NumField",
            "Field",
            "FieldByName",
            "FieldByNameFunc",
            "MapIndex",
            "Len",
            "Index",
            "IsValid",
            "IsNil",
            "CanSet",
            "IsZero",
            "Elem",
            "Set",
            "SetInt",
            "SetUint",
            "SetString",
            "SetBool",
            "MethodByName",
            "Call",
            "CallSlice",
        ],
    );
    register_type(
        root,
        "__goReflectValue",
        &[
            "Kind",
            "Name",
            "Interface",
            "Int",
            "Uint",
            "Float",
            "Bool",
            "String",
            "NumField",
            "Field",
            "FieldByName",
            "FieldByNameFunc",
            "MapIndex",
            "Len",
            "Index",
            "IsValid",
            "IsNil",
            "CanSet",
            "IsZero",
            "Elem",
            "Set",
            "SetInt",
            "SetUint",
            "SetString",
            "SetBool",
            "MethodByName",
            "Call",
            "CallSlice",
        ],
    );
}

fn emit_name(method: &str) -> Option<&'static str> {
    match method {
        "Kind" => Some("go.reflect_kind"),
        "Name" => Some("go.reflect_name"),
        "Interface" => Some("go.reflect_interface"),
        "Int" => Some("go.reflect_int"),
        "Uint" => Some("go.reflect_uint"),
        "Float" => Some("go.reflect_float"),
        "Bool" => Some("go.reflect_bool"),
        "String" => Some("go.reflect_string"),
        "NumField" => Some("go.reflect_num_field"),
        "Field" => Some("go.reflect_field"),
        "FieldByName" | "FieldByNameFunc" => Some("go.reflect_field_by_name"),
        "MapIndex" => Some("go.reflect_map_index"),
        "Len" => Some("go.reflect_len"),
        "Index" => Some("go.reflect_index"),
        "IsValid" => Some("go.reflect_is_valid"),
        "IsNil" => Some("go.reflect_is_nil"),
        "CanSet" => Some("go.reflect_can_set"),
        "IsZero" => Some("go.reflect_is_zero"),
        "Elem" => Some("go.reflect_elem"),
        "Set" => Some("go.reflect_set"),
        "SetInt" => Some("go.reflect_set_int"),
        "SetUint" => Some("go.reflect_set_uint"),
        "SetString" => Some("go.reflect_set_string"),
        "SetBool" => Some("go.reflect_set_bool"),
        _ => None,
    }
}

fn register_type(root: &mut Subtree, name: &str, methods: &[&str]) {
    let mut method_tree = Subtree::new();
    for method in methods {
        if let Some(emit) = emit_name(method) {
            method_tree.insert(
                (*method).to_string(),
                NamespaceNode::CommonEmit(emit.to_string()),
            );
        }
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
