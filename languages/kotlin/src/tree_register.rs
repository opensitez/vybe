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

        namespaces::register_namespace_tree("kotlin", NamespaceNode::Namespace(root));
    });
}
