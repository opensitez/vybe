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

fn common_method(emit: &str, min_args: u8, max_args: u8) -> NamespaceNode {
    namespaces::overloads(
        (min_args..=max_args)
            .map(|arity| (arity, common_emit(emit)))
            .collect(),
    )
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

fn kotlin_random_type() -> NamespaceNode {
    let mut statics = Subtree::new();
    statics.insert("default".to_string(), common_emit("kotlin.random_default"));

    let mut methods = Subtree::new();
    for (name, emit, min_args, max_args) in [
        ("nextint", "jvm.java.random_next_int", 0, 2),
        ("nextlong", "jvm.java.random_next_long", 0, 2),
        ("nextdouble", "jvm.java.random_next_double", 0, 2),
        ("nextfloat", "jvm.java.random_next_float", 0, 0),
        ("nextboolean", "jvm.java.random_next_boolean", 0, 0),
        ("nextbytes", "jvm.java.random_next_bytes", 1, 1),
    ] {
        methods.insert(name.to_string(), common_method(emit, min_args, max_args));
    }

    NamespaceNode::Type {
        ctor: None,
        ctor_call: Some(Box::new(common_emit("jvm.java.random_new"))),
        statics,
        methods,
        member_returns: [
            ("nextint".to_string(), "Int".to_string()),
            ("nextlong".to_string(), "Long".to_string()),
            ("nextdouble".to_string(), "Double".to_string()),
            ("nextfloat".to_string(), "Float".to_string()),
            ("nextboolean".to_string(), "Boolean".to_string()),
            ("nextbytes".to_string(), "ByteArray".to_string()),
        ]
        .into_iter()
        .collect(),
    }
}

fn kotlin_regex_type() -> NamespaceNode {
    let mut methods = Subtree::new();
    for (name, emit, min_args, max_args) in [
        ("matches", "kotlin.regex_matches", 1, 1),
        ("matchEntire", "kotlin.regex_match_entire", 1, 1),
        ("containsMatchIn", "kotlin.regex_contains", 1, 1),
        ("find", "kotlin.regex_find", 1, 2),
        ("findAll", "kotlin.regex_find_all", 1, 2),
        ("matchesAt", "kotlin.regex_matches_at", 2, 2),
        ("split", "kotlin.regex_split", 1, 2),
        ("splitToSequence", "kotlin.regex_split", 1, 2),
        ("replace", "kotlin.regex_replace", 2, 2),
        ("replaceFirst", "kotlin.regex_replace_first", 2, 2),
        ("toPattern", "kotlin.regex_to_pattern", 0, 0),
        ("pattern", "kotlin.regex_pattern", 0, 0),
    ] {
        methods.insert(name.to_ascii_lowercase(), common_method(emit, min_args, max_args));
    }

    // `Regex.escape` and `Regex.fromLiteral` are companion-object functions, so
    // they hang off the TYPE rather than an instance.
    let mut statics = Subtree::new();
    statics.insert(
        "escape".to_string(),
        common_method("kotlin.regex_escape", 1, 1),
    );
    statics.insert(
        "escapereplacement".to_string(),
        common_method("kotlin.regex_escape", 1, 1),
    );
    statics.insert(
        "fromliteral".to_string(),
        common_method("kotlin.regex_from_literal", 1, 1),
    );

    NamespaceNode::Type {
        ctor: None,
        ctor_call: Some(Box::new(common_emit("kotlin.regex_new"))),
        statics,
        methods,
        member_returns: [
            ("matches".to_string(), "Boolean".to_string()),
            (
                "matchentire".to_string(),
                "java.util.regex.Matcher".to_string(),
            ),
            ("containsmatchin".to_string(), "Boolean".to_string()),
            ("find".to_string(), "java.util.regex.Matcher".to_string()),
            ("matchesat".to_string(), "Boolean".to_string()),
            (
                "findall".to_string(),
                "List<java.util.regex.Matcher>".to_string(),
            ),
            ("split".to_string(), "List".to_string()),
            ("splittosequence".to_string(), "Sequence".to_string()),
            ("replace".to_string(), "String".to_string()),
            ("replacefirst".to_string(), "String".to_string()),
            ("topattern".to_string(), "java.util.regex.Pattern".to_string()),
            ("pattern".to_string(), "String".to_string()),
        ]
        .into_iter()
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
            ("math.log", "ecma:math", "log"),
            ("math.log10", "ecma:math", "log10"),
            ("math.hypot", "ecma:math", "hypot"),
            ("math.max", "ecma:math", "maxOf"),
            ("math.min", "ecma:math", "minOf"),
            ("math.isfinite", "ecma:number", "isFinite"),
        ] {
            insert_path(&mut root, name, namespaces::host_fn(module, func));
        }

        for (name, emit) in [
            ("math.round", "kotlin.round"),
            ("math.sign", "kotlin.sign"),
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
            ("collections.containsall", "jvm.java.list_contains_all"),
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
            ("containsall", "jvm.java.list_contains_all"),
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
                    ("add", "jvm.java.add"),
                    ("remove", "kotlin.remove_any"),
                    ("clear", "kotlin.clear_any"),
                    ("union", "kotlin.set_union"),
                    ("intersect", "kotlin.set_intersect"),
                    ("subtract", "kotlin.set_subtract"),
                    ("containsall", "jvm.java.list_contains_all"),
                    ("toset", "kotlin.to_set"),
                    ("tomutableset", "kotlin.to_set"),
                    ("tohashset", "kotlin.to_hash_set"),
                ],
                &set_returns,
            ),
        );

        insert_path(&mut root, "random.random", kotlin_random_type());
        insert_path(&mut root, "text.regex", kotlin_regex_type());
        for (name, value) in [
            ("text.regexoption.ignore_case", 2.0),
            ("text.regexoption.comments", 4.0),
            ("text.regexoption.multiline", 8.0),
            ("text.regexoption.literal", 16.0),
            ("text.regexoption.dot_matches_all", 32.0),
            ("text.regexoption.unix_lines", 1.0),
            ("text.regexoption.canon_eq", 128.0),
        ] {
            insert_path(&mut root, name, NamespaceNode::Const(Value::F64(value)));
        }
        for (name, emit, min_args, max_args) in [
            ("random.nextint", "jvm.java.random_next_int", 1, 3),
            ("random.nextlong", "jvm.java.random_next_long", 1, 3),
            ("random.nextdouble", "jvm.java.random_next_double", 1, 3),
            ("random.nextfloat", "jvm.java.random_next_float", 1, 1),
            ("random.nextboolean", "jvm.java.random_next_boolean", 1, 1),
            ("random.nextbytes", "jvm.java.random_next_bytes", 2, 2),
        ] {
            insert_path(&mut root, name, common_method(emit, min_args, max_args));
        }

        for (name, emit) in [
            ("concurrent.thread", "kotlin.thread_make"),
            ("system.measuretimemillis", "kotlin.measure_time_millis"),
            ("system.measurenanotime", "kotlin.measure_nano_time"),
            ("system.identityhashcode", "kotlin.identity_hash_code"),
        ] {
            insert_path(&mut root, name, common_emit(emit));
        }

        namespaces::register_namespace_tree("kotlin", NamespaceNode::Namespace(root));
    });
}
