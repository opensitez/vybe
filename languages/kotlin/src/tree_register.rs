//! Kotlin namespace-tree registration.
//!
//! Kotlin-owned namespaced library surface (`kotlin.*`) registers here as
//! tree leaves. The JVM platform owns `jvm.java.*`; it must not mount Kotlin
//! names or carry Kotlin package rows.

use std::sync::Once;

use vybe_runtime::Value;
use vybe_compiler::primitives::namespaces::{self, NamespaceNode, Subtree};

fn insert_path(root: &mut Subtree, path: &str, node: NamespaceNode) {
    // ⛔ The path keeps the spelling written above it. It used to be lowercased
    // here, which silently overrode every declaration in this file: a row could
    // say `math.roundToInt` and the tree would still hold `roundtoint`. That was
    // invisible while the tree folded on a miss; now that a case-sensitive
    // language asks EXACTLY, a folded key is simply unreachable.
    let mut segments: Vec<String> = path.split('.').map(|s| s.to_string()).collect();
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
        // Declared spelling — see `insert_path`.
        method_tree.insert((*name).to_string(), common_emit(emit));
    }
    NamespaceNode::Type {
        ctor: None,
        ctor_call: None,
        statics: Subtree::new(),
        methods: method_tree,
        member_returns: returns
            .iter()
            .map(|(name, ty)| ((*name).to_string(), (*ty).to_string()))
            .collect(),
    }
}

fn kotlin_random_type() -> NamespaceNode {
    let mut statics = Subtree::new();
    // `Random.Default` — Kotlin spells the companion with a capital D.
    statics.insert("Default".to_string(), common_emit("kotlin.random_default"));

    let mut methods = Subtree::new();
    for (name, emit, min_args, max_args) in [
        ("nextInt", "jvm.java.random_next_int", 0, 2),
        ("nextLong", "jvm.java.random_next_long", 0, 2),
        ("nextDouble", "jvm.java.random_next_double", 0, 2),
        ("nextFloat", "jvm.java.random_next_float", 0, 0),
        ("nextBoolean", "jvm.java.random_next_boolean", 0, 0),
        ("nextBytes", "jvm.java.random_next_bytes", 1, 1),
    ] {
        methods.insert(name.to_string(), common_method(emit, min_args, max_args));
    }

    NamespaceNode::Type {
        ctor: None,
        ctor_call: Some(Box::new(common_emit("jvm.java.random_new"))),
        statics,
        methods,
        member_returns: [
            ("nextInt".to_string(), "Int".to_string()),
            ("nextLong".to_string(), "Long".to_string()),
            ("nextDouble".to_string(), "Double".to_string()),
            ("nextFloat".to_string(), "Float".to_string()),
            ("nextBoolean".to_string(), "Boolean".to_string()),
            ("nextBytes".to_string(), "ByteArray".to_string()),
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
        // Declared spelling — see `insert_path`.
        methods.insert(name.to_string(), common_method(emit, min_args, max_args));
    }

    // `Regex.escape` and `Regex.fromLiteral` are companion-object functions, so
    // they hang off the TYPE rather than an instance.
    let mut statics = Subtree::new();
    statics.insert(
        "escape".to_string(),
        common_method("kotlin.regex_escape", 1, 1),
    );
    statics.insert(
        "escapeReplacement".to_string(),
        common_method("kotlin.regex_escape", 1, 1),
    );
    statics.insert(
        "fromLiteral".to_string(),
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
                "matchEntire".to_string(),
                "java.util.regex.Matcher".to_string(),
            ),
            ("containsMatchIn".to_string(), "Boolean".to_string()),
            ("find".to_string(), "java.util.regex.Matcher".to_string()),
            ("matchesAt".to_string(), "Boolean".to_string()),
            (
                "findAll".to_string(),
                "List<java.util.regex.Matcher>".to_string(),
            ),
            ("split".to_string(), "List".to_string()),
            ("splitToSequence".to_string(), "Sequence".to_string()),
            ("replace".to_string(), "String".to_string()),
            ("replaceFirst".to_string(), "String".to_string()),
            ("toPattern".to_string(), "java.util.regex.Pattern".to_string()),
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
            ("math.absoluteValue", "ecma:math", "abs"),
            ("math.sqrt", "ecma:math", "sqrt"),
            ("math.pow", "ecma:math", "pow"),
            ("math.roundToInt", "ecma:math", "round"),
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
            ("math.isFinite", "ecma:number", "isFinite"),
        ] {
            insert_path(&mut root, name, namespaces::host_fn(module, func));
        }

        for (name, emit) in [
            ("math.round", "kotlin.round"),
            ("math.sign", "kotlin.sign"),
            ("math.ulp", "jvm.java.math_ulp"),
            ("math.nextTowards", "jvm.java.math_next_after"),
            ("math.nextUp", "jvm.java.math_next_up"),
            ("math.nextDown", "jvm.java.math_next_down"),
        ] {
            insert_path(&mut root, name, common_emit(emit));
        }

        insert_path(
            &mut root,
            "math.PI",
            NamespaceNode::Const(Value::F64(std::f64::consts::PI)),
        );
        insert_path(
            &mut root,
            "math.E",
            NamespaceNode::Const(Value::F64(std::f64::consts::E)),
        );

        for (name, emit) in [
            ("collections.setOf", "kotlin.set_literal"),
            ("collections.mutableSetOf", "kotlin.set_literal"),
            ("collections.emptySet", "kotlin.set_literal"),
            ("collections.buildSet", "kotlin.set_literal"),
            ("collections.linkedSetOf", "kotlin.set_literal"),
            ("collections.hashSetOf", "kotlin.hash_set_literal"),
            ("collections.union", "kotlin.set_union"),
            ("collections.intersect", "kotlin.set_intersect"),
            ("collections.subtract", "kotlin.set_subtract"),
            ("collections.containsAll", "jvm.java.list_contains_all"),
        ] {
            insert_path(&mut root, name, common_emit(emit));
        }

        for (name, target) in [
            ("collections.ArrayList", "jvm.java.util.ArrayList"),
            ("collections.HashMap", "jvm.java.util.HashMap"),
            ("collections.HashSet", "jvm.java.util.HashSet"),
            ("collections.LinkedHashMap", "jvm.java.util.LinkedHashMap"),
            ("collections.LinkedHashSet", "jvm.java.util.LinkedHashSet"),
            ("text.StringBuilder", "jvm.java.lang.StringBuilder"),
        ] {
            insert_path(&mut root, name, NamespaceNode::Alias(target.to_string()));
        }

        let set_methods = [
            ("union", "kotlin.set_union"),
            ("intersect", "kotlin.set_intersect"),
            ("subtract", "kotlin.set_subtract"),
            ("containsAll", "jvm.java.list_contains_all"),
            ("toSet", "kotlin.to_set"),
            ("toMutableSet", "kotlin.to_set"),
            ("toHashSet", "kotlin.to_hash_set"),
        ];
        let set_returns = [
            ("union", "Set"),
            ("intersect", "Set"),
            ("subtract", "Set"),
            ("toSet", "Set"),
            ("toMutableSet", "Set"),
            ("toHashSet", "Set"),
        ];
        insert_path(
            &mut root,
            "collections.Set",
            kotlin_collection_type(&set_methods, &set_returns),
        );
        insert_path(
            &mut root,
            "collections.MutableSet",
            kotlin_collection_type(
                &[
                    ("add", "jvm.java.add"),
                    ("remove", "kotlin.remove_any"),
                    ("clear", "kotlin.clear_any"),
                    ("union", "kotlin.set_union"),
                    ("intersect", "kotlin.set_intersect"),
                    ("subtract", "kotlin.set_subtract"),
                    ("containsAll", "jvm.java.list_contains_all"),
                    ("toSet", "kotlin.to_set"),
                    ("toMutableSet", "kotlin.to_set"),
                    ("toHashSet", "kotlin.to_hash_set"),
                ],
                &set_returns,
            ),
        );

        insert_path(&mut root, "random.Random", kotlin_random_type());
        insert_path(&mut root, "text.Regex", kotlin_regex_type());
        for (name, value) in [
            ("text.RegexOption.IGNORE_CASE", 2.0),
            ("text.RegexOption.COMMENTS", 4.0),
            ("text.RegexOption.MULTILINE", 8.0),
            ("text.RegexOption.LITERAL", 16.0),
            ("text.RegexOption.DOT_MATCHES_ALL", 32.0),
            ("text.RegexOption.UNIX_LINES", 1.0),
            ("text.RegexOption.CANON_EQ", 128.0),
        ] {
            insert_path(&mut root, name, NamespaceNode::Const(Value::F64(value)));
        }
        for (name, emit, min_args, max_args) in [
            ("random.nextInt", "jvm.java.random_next_int", 1, 3),
            ("random.nextLong", "jvm.java.random_next_long", 1, 3),
            ("random.nextDouble", "jvm.java.random_next_double", 1, 3),
            ("random.nextFloat", "jvm.java.random_next_float", 1, 1),
            ("random.nextBoolean", "jvm.java.random_next_boolean", 1, 1),
            ("random.nextBytes", "jvm.java.random_next_bytes", 2, 2),
        ] {
            insert_path(&mut root, name, common_method(emit, min_args, max_args));
        }

        for (name, emit) in [
            ("concurrent.thread", "kotlin.thread_make"),
            ("system.measureTimeMillis", "kotlin.measure_time_millis"),
            ("system.measureNanoTime", "kotlin.measure_nano_time"),
            ("system.identityHashCode", "kotlin.identity_hash_code"),
        ] {
            insert_path(&mut root, name, common_emit(emit));
        }

        namespaces::register_namespace_tree("kotlin", NamespaceNode::Namespace(root));
    });
}
