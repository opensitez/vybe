//! Kotlin namespace-tree registration.
//!
//! Kotlin-owned namespaced library surface (`kotlin.*`) registers here as
//! tree leaves. The JVM platform owns `jvm.java.*`; it must not mount Kotlin
//! names or carry Kotlin package rows.

use std::sync::Once;

use vybe_runtime::Value;
use vybe_runtime::namespaces::{self, NamespaceNode, Subtree};

fn insert_path(root: &mut Subtree, path: &str, node: NamespaceNode) {
    let mut segments: Vec<String> = path.split('.').map(|s| s.to_ascii_lowercase()).collect();
    let leaf = segments.pop().expect("non-empty path");
    let mut cursor = root;
    for seg in segments {
        let entry = cursor
            .entry(seg)
            .or_insert_with(|| NamespaceNode::Namespace(Subtree::new()));
        let NamespaceNode::Namespace(children) = entry else {
            return;
        };
        cursor = children;
    }
    cursor.entry(leaf).or_insert(node);
}

fn common_emit(name: &str) -> NamespaceNode {
    NamespaceNode::CommonEmit(name.to_string())
}

fn kotlin_collection_type(methods: &[(&str, &str)], returns: &[(&str, &str)]) -> NamespaceNode {
    let mut method_tree = Subtree::new();
    for (name, emit) in methods {
        method_tree.insert(name.to_ascii_lowercase(), common_emit(emit));
    }
    NamespaceNode::Type {
        ctor: None,
        ctor_call: None,
        statics: Subtree::new(),
        methods: method_tree,
        member_returns: returns
            .iter()
            .map(|(name, ty)| (name.to_ascii_lowercase(), (*ty).to_string()))
            .collect(),
    }
}

pub fn register_namespace_tree() {
    static ONCE: Once = Once::new();
    ONCE.call_once(|| {
        let mut root = Subtree::new();

        for (name, module, func) in [
            ("math.abs", "ecma:math", "abs"),
            ("math.absolutevalue", "ecma:math", "abs"),
            ("math.sqrt", "ecma:math", "sqrt"),
            ("math.pow", "ecma:math", "pow"),
            ("math.round", "ecma:math", "round"),
            ("math.roundtoint", "ecma:math", "round"),
            ("math.ceil", "ecma:math", "ceil"),
            ("math.floor", "ecma:math", "floor"),
            ("math.sin", "ecma:math", "sin"),
            ("math.cos", "ecma:math", "cos"),
            ("math.tan", "ecma:math", "tan"),
            ("math.asin", "ecma:math", "asin"),
            ("math.acos", "ecma:math", "acos"),
            ("math.atan", "ecma:math", "atan"),
            ("math.atan2", "ecma:math", "atan2"),
            ("math.sinh", "ecma:math", "sinh"),
            ("math.cosh", "ecma:math", "cosh"),
            ("math.tanh", "ecma:math", "tanh"),
            ("math.exp", "ecma:math", "exp"),
            ("math.ln", "ecma:math", "log"),
            ("math.log10", "ecma:math", "log10"),
            ("math.hypot", "ecma:math", "hypot"),
            ("math.max", "ecma:math", "maxOf"),
            ("math.min", "ecma:math", "minOf"),
            ("math.isfinite", "ecma:number", "isFinite"),
        ] {
            insert_path(&mut root, name, namespaces::host_fn(module, func));
        }

        for (name, emit) in [
            ("math.sign", "jvm.java.signum"),
            ("math.ulp", "jvm.java.math_ulp"),
            ("math.nextafter", "jvm.java.math_next_after"),
            ("math.nextup", "jvm.java.math_next_up"),
            ("math.nextdown", "jvm.java.math_next_down"),
        ] {
            insert_path(&mut root, name, common_emit(emit));
        }

        insert_path(
            &mut root,
            "math.pi",
            NamespaceNode::Const(Value::F64(std::f64::consts::PI)),
        );
        insert_path(
            &mut root,
            "math.e",
            NamespaceNode::Const(Value::F64(std::f64::consts::E)),
        );

        for (name, emit) in [
            ("collections.setof", "kotlin.set_literal"),
            ("collections.mutablesetof", "kotlin.set_literal"),
            ("collections.emptyset", "kotlin.set_literal"),
            ("collections.buildset", "kotlin.set_literal"),
            ("collections.linkedsetof", "kotlin.set_literal"),
            ("collections.hashsetof", "kotlin.hash_set_literal"),
            ("collections.union", "kotlin.set_union"),
            ("collections.intersect", "kotlin.set_intersect"),
            ("collections.subtract", "kotlin.set_subtract"),
            ("collections.containsall", "kotlin.contains_all"),
        ] {
            insert_path(&mut root, name, common_emit(emit));
        }

        for (name, target) in [
            ("collections.arraylist", "jvm.java.util.arraylist"),
            ("collections.hashmap", "jvm.java.util.hashmap"),
            ("collections.hashset", "jvm.java.util.hashset"),
            ("collections.linkedhashmap", "jvm.java.util.linkedhashmap"),
            ("collections.linkedhashset", "jvm.java.util.linkedhashset"),
            ("text.stringbuilder", "jvm.java.lang.stringbuilder"),
        ] {
            insert_path(&mut root, name, NamespaceNode::Alias(target.to_string()));
        }

        let set_methods = [
            ("union", "kotlin.set_union"),
            ("intersect", "kotlin.set_intersect"),
            ("subtract", "kotlin.set_subtract"),
            ("containsall", "kotlin.contains_all"),
            ("toset", "kotlin.to_set"),
            ("tomutableset", "kotlin.to_set"),
            ("tohashset", "kotlin.to_hash_set"),
        ];
        let set_returns = [
            ("union", "Set"),
            ("intersect", "Set"),
            ("subtract", "Set"),
            ("toset", "Set"),
            ("tomutableset", "Set"),
            ("tohashset", "Set"),
        ];
        insert_path(
            &mut root,
            "collections.set",
            kotlin_collection_type(&set_methods, &set_returns),
        );
        insert_path(
            &mut root,
            "collections.mutableset",
            kotlin_collection_type(
                &[
                    ("add", "kotlin.add"),
                    ("remove", "kotlin.remove_any"),
                    ("clear", "kotlin.clear_any"),
                    ("union", "kotlin.set_union"),
                    ("intersect", "kotlin.set_intersect"),
                    ("subtract", "kotlin.set_subtract"),
                    ("containsall", "kotlin.contains_all"),
                    ("toset", "kotlin.to_set"),
                    ("tomutableset", "kotlin.to_set"),
                    ("tohashset", "kotlin.to_hash_set"),
                ],
                &set_returns,
            ),
        );

        namespaces::register_namespace_tree("kotlin", NamespaceNode::Namespace(root));
    });
}
