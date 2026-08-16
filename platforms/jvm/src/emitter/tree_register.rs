//! `java.*` namespace-tree registration — the JDK as a PLATFORM.
//!
//! Mirrors the dotnet registrar: this crate contributes DATA — explicit
//! `java.*` tree leaves — to the shared namespace tree. Resolution
//! LOGIC lives only in the common resolver, so ANY language can walk
//! `java.util.objects.equals` by declaring tree data in its profile, exactly
//! as csharp/vb reach `dotnet.*` with zero `System.*` entries of their own.
//!
//! This used to live in `languages/java` and register through the LANGUAGE
//! hook, which made the JDK the property of one frontend.
//!
//! Leaf rules (dotnet template):
//! - Java package-surface common emits register as statics on type nodes
//!   (`java.util.Objects.equals` → member `equals` on `java.util.Objects`),
//!   even when the actual common op is a shared category such as
//!   `object.equals`;
//! - host-backed statics register as `Fn` leaves under the same tree;
//! - opcode/intrinsic/print builtins have no process-global target to
//!   point at — skipped.

use std::collections::BTreeMap;
use std::sync::Once;

use vybe_runtime::Value;
use vybe_runtime::namespaces::{self, NamespaceNode, Subtree};

/// Insert `node` at the dotted `path` under `root`, creating interior
/// namespaces as needed. Keys are lowercase-canonical.
fn insert_path(root: &mut Subtree, path: &str, node: NamespaceNode) {
    let mut segments: Vec<&str> = path.split('.').collect();
    let leaf = segments.pop().expect("non-empty path");
    let mut cursor = root;
    for seg in segments {
        let entry = cursor
            .entry(seg.to_string())
            .or_insert_with(|| NamespaceNode::Namespace(Subtree::new()));
        cursor = match entry {
            NamespaceNode::Namespace(children) => children,
            // A TYPE used as a namespace — `java.util.BitSet.valueOf` descends
            // THROUGH the `BitSet` type node to its statics. `resolve_segments`
            // already walks it that way; the insert was missing the symmetric
            // case, so every static under a registered type was dropped
            // SILENTLY (the arm below returns without a word). Types register
            // first, so this hit `BitSet.valueOf` the moment `BitSet` gained a
            // `JAVA_TYPES` row.
            NamespaceNode::Type { statics, .. } => statics,
            _ => return, // genuine leaf/namespace collision: first wins
        };
    }
    cursor.entry(leaf.to_string()).or_insert(node);
}

/// Make the node at `path` a `Type`, preserving whatever is already there.
///
/// `type_scopes` resolution (`find_type_node`) only descends through
/// `Namespace` nodes and matches `Type` nodes, so a class whose members live
/// under a plain `Namespace` is invisible to it. Promoting keeps the children
/// as the type's `statics`, which is exactly where a static member belongs.
fn ensure_type_node(root: &mut Subtree, path: &str) {
    let mut segments: Vec<&str> = path.split('.').collect();
    let leaf = segments.pop().expect("non-empty path");
    let mut cursor = root;
    for seg in segments {
        let entry = cursor
            .entry(seg.to_string())
            .or_insert_with(|| NamespaceNode::Namespace(Subtree::new()));
        cursor = match entry {
            NamespaceNode::Namespace(children) => children,
            NamespaceNode::Type { statics, .. } => statics,
            _ => return,
        };
    }
    // A type with no statics (`DayOfWeek` — instance surface only) has no
    // node yet: CREATE it, or the merge helpers silently no-op and every
    // instance member registered for it is dropped.
    let existing = cursor
        .entry(leaf.to_string())
        .or_insert_with(|| NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics: Subtree::new(),
            methods: Subtree::new(),
            member_returns: Default::default(),
        });
    if let NamespaceNode::Namespace(children) = existing {
        let statics = std::mem::take(children);
        *existing = NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics,
            methods: Subtree::new(),
            member_returns: Default::default(),
        };
    }
}

fn merge_type_methods(root: &mut Subtree, path: &str, new_methods: Subtree) {
    let mut segments: Vec<&str> = path.split('.').collect();
    let leaf = segments.pop().expect("non-empty path");
    let mut cursor = root;
    for seg in segments {
        let entry = cursor
            .entry(seg.to_string())
            .or_insert_with(|| NamespaceNode::Namespace(Subtree::new()));
        cursor = match entry {
            NamespaceNode::Namespace(children) => children,
            NamespaceNode::Type { statics, .. } => statics,
            _ => return,
        };
    }
    let Some(NamespaceNode::Type { methods, .. }) = cursor.get_mut(leaf) else {
        return;
    };
    for (name, node) in new_methods {
        methods.entry(name).or_insert(node);
    }
}

fn merge_type_member_returns(root: &mut Subtree, path: &str, returns: &[(&str, &str)]) {
    let mut segments: Vec<&str> = path.split('.').collect();
    let leaf = segments.pop().expect("non-empty path");
    let mut cursor = root;
    for seg in segments {
        let entry = cursor
            .entry(seg.to_string())
            .or_insert_with(|| NamespaceNode::Namespace(Subtree::new()));
        cursor = match entry {
            NamespaceNode::Namespace(children) => children,
            NamespaceNode::Type { statics, .. } => statics,
            _ => return,
        };
    }
    let Some(NamespaceNode::Type { member_returns, .. }) = cursor.get_mut(leaf) else {
        return;
    };
    for (member, ty) in returns {
        member_returns.insert(member.to_lowercase(), (*ty).to_string());
    }
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

fn insert_common_static(root: &mut Subtree, type_path: &str, member: &str, emit: &str) {
    insert_path(root, &format!("{type_path}.{member}"), common_emit(emit));
    ensure_type_node(root, type_path);
}

fn insert_host_static(root: &mut Subtree, type_path: &str, member: &str, module: &str, func: &str) {
    insert_path(
        root,
        &format!("{type_path}.{member}"),
        namespaces::host_fn(module, func),
    );
    ensure_type_node(root, type_path);
}

#[derive(Clone, Copy)]
enum JavaConst {
    Bool(bool),
    Float(f64),
    Str(&'static str),
}

fn insert_java_namespace_constants(root: &mut Subtree) {
    const SPECS: &[(&str, JavaConst)] = &[
        ("java.lang.Math.PI", JavaConst::Float(3.141592653589793)),
        ("java.lang.Math.E", JavaConst::Float(2.718281828459045)),
        (
            "java.lang.StrictMath.PI",
            JavaConst::Float(3.141592653589793),
        ),
        (
            "java.lang.StrictMath.E",
            JavaConst::Float(2.718281828459045),
        ),
        (
            "java.lang.Double.MAX_VALUE",
            JavaConst::Float(1.7976931348623157e308),
        ),
        ("java.lang.Double.MIN_VALUE", JavaConst::Float(5e-324)),
        ("java.lang.Double.MIN_EXPONENT", JavaConst::Float(-1022.0)),
        (
            "java.lang.Integer.MAX_VALUE",
            JavaConst::Float(2147483647.0),
        ),
        (
            "java.lang.Integer.MIN_VALUE",
            JavaConst::Float(-2147483648.0),
        ),
        ("java.lang.Thread.MIN_PRIORITY", JavaConst::Float(1.0)),
        ("java.lang.Thread.NORM_PRIORITY", JavaConst::Float(5.0)),
        ("java.lang.Thread.MAX_PRIORITY", JavaConst::Float(10.0)),
        (
            "java.lang.Long.MAX_VALUE",
            JavaConst::Float(9.223372036854776e18),
        ),
        (
            "java.lang.Long.MIN_VALUE",
            JavaConst::Float(-9.223372036854776e18),
        ),
        ("java.lang.Double.NaN", JavaConst::Float(f64::NAN)),
        (
            "java.lang.Double.POSITIVE_INFINITY",
            JavaConst::Float(f64::INFINITY),
        ),
        (
            "java.lang.Double.NEGATIVE_INFINITY",
            JavaConst::Float(f64::NEG_INFINITY),
        ),
        ("java.lang.Boolean.TRUE", JavaConst::Bool(true)),
        ("java.lang.Boolean.FALSE", JavaConst::Bool(false)),
        (
            "java.lang.Character.UPPERCASE_LETTER",
            JavaConst::Float(1.0),
        ),
        ("java.time.ZoneOffset.UTC", JavaConst::Str("Z")),
        (
            "java.time.ZoneId.SHORT_IDS",
            JavaConst::Str("common:jvm.java.zone_id_short_ids"),
        ),
        (
            "java.time.temporal.ChronoUnit.SECONDS",
            JavaConst::Str("SECONDS"),
        ),
        (
            "java.time.temporal.ChronoUnit.MILLIS",
            JavaConst::Str("MILLIS"),
        ),
        (
            "java.time.temporal.ChronoUnit.MINUTES",
            JavaConst::Str("MINUTES"),
        ),
        (
            "java.time.temporal.ChronoUnit.HOURS",
            JavaConst::Str("HOURS"),
        ),
        ("java.time.temporal.ChronoUnit.DAYS", JavaConst::Str("DAYS")),
        (
            "java.time.temporal.ChronoUnit.WEEKS",
            JavaConst::Str("WEEKS"),
        ),
        (
            "java.time.temporal.ChronoUnit.MONTHS",
            JavaConst::Str("MONTHS"),
        ),
        (
            "java.time.temporal.ChronoField.DAY_OF_MONTH",
            JavaConst::Str("DAY_OF_MONTH"),
        ),
        (
            "java.time.temporal.ChronoField.HOUR_OF_DAY",
            JavaConst::Str("HOUR_OF_DAY"),
        ),
        (
            "java.time.temporal.IsoFields.WEEK_OF_WEEK_BASED_YEAR",
            JavaConst::Str("WEEK_OF_WEEK_BASED_YEAR"),
        ),
        (
            "java.time.temporal.IsoFields.WEEK_BASED_YEAR",
            JavaConst::Str("WEEK_BASED_YEAR"),
        ),
        (
            "java.time.format.DateTimeFormatter.ISO_OFFSET_DATE_TIME",
            JavaConst::Str("ISO_OFFSET_DATE_TIME"),
        ),
        (
            "java.time.format.DateTimeFormatter.ISO_ZONED_DATE_TIME",
            JavaConst::Str("ISO_ZONED_DATE_TIME"),
        ),
        ("java.util.Locale.FRANCE", JavaConst::Str("FR")),
        ("java.util.Locale.GERMANY", JavaConst::Str("DE")),
        ("java.util.Locale.ITALY", JavaConst::Str("IT")),
        ("java.util.Locale.US", JavaConst::Str("US")),
        ("java.util.Locale.UK", JavaConst::Str("UK")),
        ("java.util.Locale.JAPAN", JavaConst::Str("JP")),
        ("java.util.Locale.CANADA", JavaConst::Str("CA")),
        ("java.util.Locale.CANADA_FRENCH", JavaConst::Str("FR_CA")),
    ];

    for (name, value) in SPECS {
        let node = match *value {
            JavaConst::Bool(v) => NamespaceNode::Const(Value::Bool(v)),
            JavaConst::Float(v) => NamespaceNode::Const(Value::F64(v)),
            JavaConst::Str(v) => NamespaceNode::Const(Value::String(v.into())),
        };
        let key = name.to_lowercase();
        let path = key.strip_prefix("java.").unwrap_or(key.as_str());
        insert_path(root, path, node);
    }
}

fn java_type_ctor_target(qualified: &str) -> Option<NamespaceNode> {
    let emit = match qualified {
        "java.util.TreeSet" => "jvm.java.sorted_set_new",
        "java.util.ArrayList"
        | "java.util.LinkedList"
        | "java.util.ArrayDeque"
        | "java.util.Stack" => "jvm.java.mutable_list_new",
        "java.util.concurrent.CopyOnWriteArrayList" => "jvm.java.copy_on_write_list_new",
        "java.util.concurrent.LinkedBlockingQueue" => "jvm.java.linked_blocking_queue_new",
        "java.util.Vector" => "jvm.java.vector_new",
        "java.util.PriorityQueue" => "jvm.java.priority_queue_new",
        "java.util.Comparator" => "jvm.java.collection_passthrough_new",
        "java.util.HashSet" | "java.util.LinkedHashSet" => "jvm.java.hash_set_new",
        "java.util.TreeMap" => "jvm.java.sorted_map_new",
        "java.util.HashMap"
        | "java.util.WeakHashMap"
        | "java.util.Hashtable"
        | "java.util.Properties"
        | "java.lang.OutOfMemoryError" => "jvm.java.hash_map_new",
        "java.util.concurrent.ConcurrentHashMap" => "jvm.java.concurrent_hash_map_new",
        "java.util.IdentityHashMap" => "jvm.java.identity_hash_map_new",
        "java.util.LinkedHashMap" => "jvm.java.linked_hash_map_new",
        "java.math.BigInteger" => "jvm.java.bigint_new",
        "java.util.BitSet" => "jvm.java.bitset_new",
        "java.util.UUID" => "jvm.java.uuid_new",
        "java.util.Random" | "java.util.SplittableRandom" => "jvm.java.random_new",
        "java.util.regex.Pattern" => "jvm.java.regex_pattern_compile",
        "java.lang.StringBuilder" | "java.lang.StringBuffer" => "jvm.java.stringbuilder_new",
        "java.util.StringTokenizer" => "jvm.java.stringtokenizer_new",
        "java.lang.Object" => "jvm.java.hash_map_new",
        "java.io.ByteArrayOutputStream" => "jvm.java.io_byte_array_output_stream_new",
        "java.io.ByteArrayInputStream" => "jvm.java.io_byte_array_input_stream_new",
        "java.io.File" => "jvm.java.io_file_new",
        "java.io.PrintWriter" | "java.io.PrintStream" => "jvm.java.io_print_writer_new",
        "java.io.OutputStreamWriter" | "java.io.BufferedWriter" | "java.io.FilterWriter" => {
            "jvm.java.io_passthrough_new"
        }
        "java.io.StringWriter" | "java.io.CharArrayWriter" => "jvm.java.io_string_writer_new",
        "java.io.StringReader" | "java.io.CharArrayReader" => "jvm.java.io_string_reader_new",
        "java.io.InputStreamReader"
        | "java.io.BufferedReader"
        | "java.io.BufferedInputStream"
        | "java.io.FilterInputStream"
        | "java.io.PushbackInputStream"
        | "java.io.LineNumberReader"
        | "java.io.DataInputStream" => "jvm.java.io_passthrough_new",
        "java.io.SequenceInputStream" => "jvm.java.io_sequence_input_stream_new",
        "java.io.DataOutputStream" => "jvm.java.io_passthrough_new",
        _ => return None,
    };
    Some(common_emit(emit))
}

fn insert_java_lang_system(root: &mut Subtree) {
    let mut statics = Subtree::new();
    statics.insert(
        "getproperty".to_string(),
        namespaces::overloads(vec![
            (1, common_emit("jvm.java.lang.system_get_property")),
            (2, common_emit("jvm.java.lang.system_get_property")),
        ]),
    );
    statics.insert(
        "currenttimemillis".to_string(),
        common_emit("jvm.java.current_time_millis"),
    );
    statics.insert(
        "nanotime".to_string(),
        common_emit("jvm.java.current_time_millis"),
    );
    insert_path(
        root,
        "lang.system",
        NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics,
            methods: Subtree::new(),
            member_returns: Default::default(),
        },
    );
}

fn insert_java_util_core_statics(root: &mut Subtree) {
    for (member, emit) in [
        ("equals", "object.equals"),
        ("hash", "object.hash"),
        ("hashcode", "object.hash_code"),
        ("isnull", "object.is_null"),
        ("nonnull", "object.non_null"),
        ("compare", "object.compare"),
        ("tostring", "object.to_string_or"),
    ] {
        insert_common_static(root, "util.objects", member, emit);
    }

    insert_common_static(root, "util.uuid", "fromstring", "jvm.java.uuid_from_string");
    insert_host_static(root, "util.uuid", "randomuuid", "web:crypto", "randomUUID");
    insert_common_static(
        root,
        "util.uuid",
        "nameuuidfrombytes",
        "jvm.java.uuid_name_from_bytes",
    );
    insert_common_static(root, "util.bitset", "valueof", "jvm.java.bitset_value_of");
    insert_common_static(
        root,
        "util.concurrent.concurrenthashmap",
        "newkeyset",
        "jvm.java.hash_set_new",
    );
}

fn insert_java_lang_core_statics(root: &mut Subtree) {
    for (type_path, member, emit) in [
        ("lang.integer", "parseint", "jvm.java.parse_int"),
        ("lang.integer", "valueof", "jvm.java.parse_int"),
        ("lang.integer", "compare", "jvm.java.compare"),
        ("lang.long", "parselong", "jvm.java.parse_int"),
        ("lang.long", "valueof", "jvm.java.parse_int"),
        ("lang.long", "compare", "jvm.java.compare"),
        ("lang.double", "compare", "jvm.java.compare"),
        (
            "lang.integer",
            "tobinarystring",
            "jvm.java.to_binary_string",
        ),
        ("lang.integer", "tohexstring", "jvm.java.to_hex_string"),
        ("lang.integer", "bitcount", "jvm.java.int_bit_count"),
        (
            "lang.integer",
            "compareunsigned",
            "jvm.java.int_compare_unsigned",
        ),
        (
            "lang.integer",
            "numberofleadingzeros",
            "jvm.java.int_leading_zeros",
        ),
        (
            "lang.integer",
            "numberoftrailingzeros",
            "jvm.java.int_trailing_zeros",
        ),
        ("lang.integer", "rotateleft", "jvm.java.int_rotate_left"),
        ("lang.integer", "rotateright", "jvm.java.int_rotate_right"),
        (
            "lang.integer",
            "lowestonebit",
            "jvm.java.int_lowest_one_bit",
        ),
        (
            "lang.integer",
            "highestonebit",
            "jvm.java.int_highest_one_bit",
        ),
        ("lang.integer", "tooctalstring", "jvm.java.to_octal_string"),
        ("lang.double", "isinfinite", "jvm.java.is_infinite"),
        ("lang.math", "signum", "jvm.java.signum"),
        ("lang.math", "scalb", "jvm.java.math_scalb"),
        ("lang.math", "ulp", "jvm.java.math_ulp"),
        ("lang.math", "getexponent", "jvm.java.math_get_exponent"),
        ("lang.math", "copysign", "jvm.java.math_copy_sign"),
        ("lang.math", "nextafter", "jvm.java.math_next_after"),
        ("lang.math", "nextup", "jvm.java.math_next_up"),
        ("lang.math", "nextdown", "jvm.java.math_next_down"),
        ("lang.math", "fma", "jvm.java.math_fma"),
        ("lang.math", "expm1", "jvm.java.math_expm1"),
        ("lang.math", "log1p", "jvm.java.math_log1p"),
        ("lang.math", "todegrees", "jvm.java.math_to_degrees"),
        ("lang.math", "toradians", "jvm.java.math_to_radians"),
        ("lang.math", "ieeeremainder", "jvm.java.math_ieee_remainder"),
        ("lang.math", "addexact", "jvm.java.math_add_exact"),
        ("lang.math", "subtractexact", "jvm.java.math_subtract_exact"),
        ("lang.math", "multiplyexact", "jvm.java.math_multiply_exact"),
        (
            "lang.math",
            "incrementexact",
            "jvm.java.math_increment_exact",
        ),
        (
            "lang.math",
            "decrementexact",
            "jvm.java.math_decrement_exact",
        ),
        ("lang.math", "negateexact", "jvm.java.math_negate_exact"),
        ("lang.math", "floordiv", "jvm.java.floor_div"),
        ("lang.math", "floormod", "jvm.java.floor_mod"),
        ("lang.strictmath", "signum", "jvm.java.signum"),
        ("lang.strictmath", "scalb", "jvm.java.math_scalb"),
        ("lang.strictmath", "ulp", "jvm.java.math_ulp"),
        (
            "lang.strictmath",
            "getexponent",
            "jvm.java.math_get_exponent",
        ),
        ("lang.strictmath", "copysign", "jvm.java.math_copy_sign"),
        ("lang.strictmath", "nextafter", "jvm.java.math_next_after"),
        ("lang.strictmath", "nextup", "jvm.java.math_next_up"),
        ("lang.strictmath", "nextdown", "jvm.java.math_next_down"),
        ("lang.strictmath", "fma", "jvm.java.math_fma"),
        ("lang.strictmath", "expm1", "jvm.java.math_expm1"),
        ("lang.strictmath", "log1p", "jvm.java.math_log1p"),
        ("lang.strictmath", "todegrees", "jvm.java.math_to_degrees"),
        ("lang.strictmath", "toradians", "jvm.java.math_to_radians"),
        (
            "lang.strictmath",
            "ieeeremainder",
            "jvm.java.math_ieee_remainder",
        ),
        ("lang.character", "isdigit", "jvm.java.char_is_digit"),
        ("lang.character", "isletter", "jvm.java.char_is_letter"),
        (
            "lang.character",
            "isletterordigit",
            "jvm.java.char_is_alnum",
        ),
        ("lang.character", "isuppercase", "jvm.java.char_is_upper"),
        ("lang.character", "islowercase", "jvm.java.char_is_lower"),
        ("lang.character", "iswhitespace", "jvm.java.char_is_space"),
        ("lang.character", "touppercase", "jvm.java.char_to_upper"),
        ("lang.character", "tolowercase", "jvm.java.char_to_lower"),
        ("lang.character", "getnumericvalue", "jvm.java.char_numeric"),
        ("lang.character", "valueof", "jvm.java.identity"),
    ] {
        insert_common_static(root, type_path, member, emit);
    }

    for (type_path, member, module, func) in [
        ("lang.double", "parsedouble", "ecma:number", "Number"),
        ("lang.double", "valueof", "ecma:number", "Number"),
        ("lang.float", "parsefloat", "ecma:number", "Number"),
        ("lang.float", "valueof", "ecma:number", "Number"),
        ("math.biginteger", "valueof", "ecma:bigint", "BigInt"),
        ("lang.double", "isnan", "ecma:number", "isNaN"),
        ("lang.math", "abs", "ecma:math", "abs"),
        ("lang.math", "sqrt", "ecma:math", "sqrt"),
        ("lang.math", "floor", "ecma:math", "floor"),
        ("lang.math", "ceil", "ecma:math", "ceil"),
        ("lang.math", "round", "ecma:math", "round"),
        ("lang.math", "min", "ecma:math", "minOf"),
        ("lang.math", "max", "ecma:math", "maxOf"),
        ("lang.math", "rint", "ecma:math", "round"),
        ("lang.math", "pow", "ecma:math", "pow"),
        ("lang.math", "exp", "ecma:math", "exp"),
        ("lang.math", "log", "ecma:math", "log"),
        ("lang.math", "log10", "ecma:math", "log10"),
        ("lang.math", "sin", "ecma:math", "sin"),
        ("lang.math", "cos", "ecma:math", "cos"),
        ("lang.math", "tan", "ecma:math", "tan"),
        ("lang.math", "asin", "ecma:math", "asin"),
        ("lang.math", "acos", "ecma:math", "acos"),
        ("lang.math", "atan", "ecma:math", "atan"),
        ("lang.math", "atan2", "ecma:math", "atan2"),
        ("lang.math", "random", "ecma:math", "random"),
        ("lang.math", "hypot", "ecma:math", "hypot"),
        ("lang.math", "cbrt", "ecma:math", "cbrt"),
        ("lang.strictmath", "abs", "ecma:math", "abs"),
        ("lang.strictmath", "sqrt", "ecma:math", "sqrt"),
        ("lang.strictmath", "floor", "ecma:math", "floor"),
        ("lang.strictmath", "ceil", "ecma:math", "ceil"),
        ("lang.strictmath", "min", "ecma:math", "minOf"),
        ("lang.strictmath", "max", "ecma:math", "maxOf"),
        ("lang.strictmath", "rint", "ecma:math", "round"),
        ("lang.strictmath", "pow", "ecma:math", "pow"),
        ("lang.strictmath", "exp", "ecma:math", "exp"),
        ("lang.strictmath", "log", "ecma:math", "log"),
        ("lang.strictmath", "log10", "ecma:math", "log10"),
        ("lang.strictmath", "sin", "ecma:math", "sin"),
        ("lang.strictmath", "cos", "ecma:math", "cos"),
        ("lang.strictmath", "tan", "ecma:math", "tan"),
        ("lang.strictmath", "asin", "ecma:math", "asin"),
        ("lang.strictmath", "acos", "ecma:math", "acos"),
        ("lang.strictmath", "atan", "ecma:math", "atan"),
        ("lang.strictmath", "atan2", "ecma:math", "atan2"),
        ("lang.strictmath", "sinh", "ecma:math", "sinh"),
        ("lang.strictmath", "cosh", "ecma:math", "cosh"),
        ("lang.strictmath", "tanh", "ecma:math", "tanh"),
        ("lang.strictmath", "hypot", "ecma:math", "hypot"),
        ("lang.strictmath", "cbrt", "ecma:math", "cbrt"),
    ] {
        insert_host_static(root, type_path, member, module, func);
    }
}

/// `java.math.BigInteger` instance methods — TREE-bound (never a walker
/// name-table), so Java's qualified spellings and Kotlin's interop both
/// resolve them. Every target composes `ecma:bigint` in the platform's
/// biginteger adapter. `mod` maps to `rem` deliberately for now: Java's
/// `mod` is the non-negative variant, and the rem-shaped answer is what the
/// Java surface already gave.
fn insert_java_math_biginteger_methods(root: &mut Subtree) {
    const SPECS: &[(&str, &str, &str, u8, u8)] = &[
        (
            "math.biginteger",
            "tostring",
            "jvm.java.bigint_to_string",
            0,
            0,
        ),
        ("math.biginteger", "add", "jvm.java.bigint_add", 1, 1),
        ("math.biginteger", "subtract", "jvm.java.bigint_sub", 1, 1),
        ("math.biginteger", "multiply", "jvm.java.bigint_mul", 1, 1),
        ("math.biginteger", "divide", "jvm.java.bigint_div", 1, 1),
        ("math.biginteger", "remainder", "jvm.java.bigint_rem", 1, 1),
        ("math.biginteger", "mod", "jvm.java.bigint_rem", 1, 1),
        ("math.biginteger", "pow", "jvm.java.bigint_pow", 1, 1),
        ("math.biginteger", "gcd", "jvm.java.bigint_gcd", 1, 1),
        ("math.biginteger", "and", "jvm.java.bigint_and", 1, 1),
        ("math.biginteger", "or", "jvm.java.bigint_or", 1, 1),
        ("math.biginteger", "xor", "jvm.java.bigint_xor", 1, 1),
        ("math.biginteger", "equals", "jvm.java.bigint_eq", 1, 1),
        ("math.biginteger", "not", "jvm.java.bigint_not", 0, 0),
        ("math.biginteger", "negate", "jvm.java.bigint_neg", 0, 0),
        ("math.biginteger", "abs", "jvm.java.bigint_abs", 0, 0),
        ("math.biginteger", "shiftleft", "jvm.java.bigint_shl", 1, 1),
        ("math.biginteger", "shiftright", "jvm.java.bigint_shr", 1, 1),
        (
            "math.biginteger",
            "compareto",
            "jvm.java.bigint_compare_to",
            1,
            1,
        ),
        ("math.biginteger", "signum", "jvm.java.bigint_signum", 0, 0),
        ("math.biginteger", "min", "jvm.java.bigint_min", 1, 1),
        ("math.biginteger", "max", "jvm.java.bigint_max", 1, 1),
        (
            "math.biginteger",
            "testbit",
            "jvm.java.bigint_test_bit",
            1,
            1,
        ),
        (
            "math.biginteger",
            "bitlength",
            "jvm.java.bigint_bit_length",
            0,
            0,
        ),
        (
            "math.biginteger",
            "isprobableprime",
            "jvm.java.bigint_is_probable_prime",
            1,
            1,
        ),
        (
            "math.biginteger",
            "nextprobableprime",
            "jvm.java.bigint_next_probable_prime",
            0,
            0,
        ),
    ];
    let mut methods = Subtree::new();
    for (_, member, emit, min_args, max_args) in SPECS {
        methods.insert(
            (*member).to_string(),
            common_method(emit, *min_args, *max_args),
        );
    }
    ensure_type_node(root, "math.biginteger");
    merge_type_methods(root, "math.biginteger", methods);
}

/// The `java.util` collection surface.
///
/// Declarations are keyed by SIMPLE TYPE NAME, and the INTERFACES own the
/// shared members: `size`/`contains`/`clear` are declared once on `Collection`,
/// not once per concrete class. Each type then collects its own declarations
/// plus everything declared on its supertypes, folding the `ancestry` array
/// that [`JAVA_TYPES`] already carries.
///
/// This is the registration-time expansion every other adapter does — the tree
/// stays a flat lookup and inheritance is resolved HERE, by the adapter that
/// knows its own type graph. Before this, each concrete type re-declared its
/// whole inherited surface by hand, and the holes were exactly what you would
/// expect from hand-copying: `TreeSet` had no `size`/`contains`/`clear`/
/// `isEmpty`/`remove`/`iterator`, `Stack` had none of the `Vector` surface it
/// extends, `ArrayList` had no `contains`/`isEmpty`/`remove`, `TreeMap` had no
/// `isEmpty`/`clear`/`entrySet`, and `LinkedHashSet`/`LinkedHashMap` were
/// verbatim copies of their parents.
///
/// `ancestry` is the right input BECAUSE it lists interfaces alongside
/// supertypes: Java's member inheritance runs through `List`/`Collection`/`Map`,
/// so a single `parent` link could not express it.
///
/// Nearest declaration wins, so a subtype overrides — `TreeSet.add` binds
/// `sorted_add` rather than `Collection.add`, `PriorityQueue.peek` binds
/// `priority_peek` rather than `Queue.peek`, and `Stack.pop` binds
/// `collections.pop` rather than anything up its chain.
///
/// Only emits that some type already declared are promoted here; nothing new is
/// invented. The shared emits were already written to be representation-generic
/// — `emit_size`/`emit_contains`/`emit_clear`/`emit_remove` each branch on
/// `emit_is_ecma_set` at runtime — so a set-backed and a list-backed receiver
/// both answer correctly from one declaration.
fn insert_java_util_collection_methods(root: &mut Subtree) {
    const SPECS: &[(&str, &str, &str, u8, u8)] = &[
        // ── Iterable ───────────────────────────────────────────────────────
        ("Iterable", "iterator", "jvm.java.list_iterator", 0, 0),
        // ── Collection ─────────────────────────────────────────────────────
        ("Collection", "add", "jvm.java.add", 1, 1),
        ("Collection", "size", "jvm.java.size", 0, 0),
        ("Collection", "isempty", "jvm.java.is_empty", 0, 0),
        ("Collection", "contains", "jvm.java.contains", 1, 1),
        ("Collection", "clear", "jvm.java.list_clear", 0, 0),
        // `AbstractCollection.toString()` — see `emit_collection_to_string` for
        // why this cannot be a runtime probe: a `List` and a Java array are the
        // same ECMA array, and Java renders only one of them element-wise.
        ("Collection", "tostring", "jvm.java.collection_to_string", 0, 0),
        // `Collection.remove(Object)` is BY VALUE — `jvm.java.list_remove` at
        // this arity is `emit_remove_at`, i.e. by INDEX, which a `Set` has no
        // notion of. `List` overrides below with the index overload that is
        // genuinely List's.
        ("Collection", "remove", "jvm.java.list_remove_value", 1, 1),
        // ── List ───────────────────────────────────────────────────────────
        // `add(index, e)` is the List-only overload, so List widens the arity.
        ("List", "add", "jvm.java.add", 1, 2),
        // `List.remove(int)` — the index overload Java declares on `List` and
        // nowhere above it.
        ("List", "remove", "jvm.java.list_remove", 1, 1),
        ("List", "get", "jvm.java.get", 1, 1),
        ("List", "set", "jvm.java.list_set", 1, 2),
        ("List", "sort", "jvm.java.collections_sort", 0, 1),
        ("List", "listiterator", "jvm.java.list_iterator", 0, 1),
        // ── SortedSet / NavigableSet ───────────────────────────────────────
        ("SortedSet", "add", "jvm.java.sorted_add", 1, 1),
        ("SortedSet", "first", "jvm.java.sorted_first", 0, 0),
        ("SortedSet", "last", "jvm.java.sorted_last", 0, 0),
        ("NavigableSet", "ceiling", "jvm.java.sorted_ceiling", 1, 1),
        ("NavigableSet", "floor", "jvm.java.sorted_floor", 1, 1),
        ("NavigableSet", "higher", "jvm.java.sorted_higher", 1, 1),
        ("NavigableSet", "lower", "jvm.java.sorted_lower", 1, 1),
        (
            "NavigableSet",
            "descendingset",
            "jvm.java.sorted_descending_set",
            0,
            0,
        ),
        // ── Queue ──────────────────────────────────────────────────────────
        ("Queue", "offer", "collections.push", 1, 1),
        ("Queue", "poll", "jvm.java.queue_poll", 0, 0),
        ("Queue", "peek", "jvm.java.peek_first", 0, 0),
        // ── Deque ──────────────────────────────────────────────────────────
        ("Deque", "addfirst", "jvm.java.add_first", 1, 1),
        ("Deque", "addlast", "collections.push", 1, 1),
        ("Deque", "offerfirst", "jvm.java.add_first", 1, 1),
        ("Deque", "offerlast", "collections.push", 1, 1),
        ("Deque", "removefirst", "jvm.java.remove_first", 0, 0),
        ("Deque", "removelast", "collections.pop", 0, 0),
        ("Deque", "peekfirst", "jvm.java.peek_first", 0, 0),
        ("Deque", "peeklast", "jvm.java.peek_last", 0, 0),
        ("Deque", "push", "jvm.java.add_first", 1, 1),
        ("Deque", "pop", "jvm.java.poll_first", 0, 0),
        // ── Map (NOT a Collection — no `iterator`, its own `size`) ─────────
        ("Map", "put", "jvm.java.map_put", 2, 2),
        ("Map", "get", "jvm.java.map_get", 1, 1),
        ("Map", "size", "jvm.java.map_size", 0, 0),
        ("Map", "isempty", "jvm.java.map_is_empty", 0, 0),
        ("Map", "clear", "jvm.java.map_clear", 0, 0),
        ("Map", "keyset", "jvm.java.map_key_set", 0, 0),
        ("Map", "values", "jvm.java.map_values", 0, 0),
        ("Map", "entryset", "jvm.java.entry_set", 0, 0),
        // ── SortedMap / NavigableMap ───────────────────────────────────────
        ("SortedMap", "keyset", "jvm.java.sorted_map_key_set", 0, 0),
        ("SortedMap", "values", "jvm.java.sorted_map_values", 0, 0),
        (
            "SortedMap",
            "firstkey",
            "jvm.java.sorted_map_first_key",
            0,
            0,
        ),
        ("SortedMap", "lastkey", "jvm.java.sorted_map_last_key", 0, 0),
        (
            "NavigableMap",
            "higherkey",
            "jvm.java.sorted_map_higher_key",
            1,
            1,
        ),
        (
            "NavigableMap",
            "lowerkey",
            "jvm.java.sorted_map_lower_key",
            1,
            1,
        ),
        // ── Iterator ───────────────────────────────────────────────────────
        ("Iterator", "next", "jvm.java.iterator_next", 0, 0),
        ("Iterator", "hasnext", "jvm.java.iterator_has_next", 0, 0),
        ("Iterator", "remove", "jvm.java.iterator_remove", 0, 0),
        (
            "ListIterator",
            "previous",
            "jvm.java.iterator_previous",
            0,
            0,
        ),
        (
            "ListIterator",
            "hasprevious",
            "jvm.java.iterator_has_previous",
            0,
            0,
        ),
        (
            "ListIterator",
            "nextindex",
            "jvm.java.iterator_next_index",
            0,
            0,
        ),
        (
            "ListIterator",
            "previousindex",
            "jvm.java.iterator_previous_index",
            0,
            0,
        ),
        ("ListIterator", "set", "jvm.java.iterator_set", 1, 1),
        ("ListIterator", "add", "jvm.java.iterator_add", 1, 1),
        // ── Concrete classes: ONLY what they override or add ───────────────
        ("Vector", "addelement", "jvm.java.add", 1, 1),
        ("Vector", "elementat", "jvm.java.get", 1, 1),
        // `Stack` is a `Vector`, so it inherits the whole list surface; these
        // four are its own, and `pop`/`peek` work the far end from `Deque`.
        ("Stack", "push", "jvm.java.add", 1, 1),
        ("Stack", "pop", "collections.pop", 0, 0),
        ("Stack", "peek", "jvm.java.peek_last", 0, 0),
        ("Stack", "empty", "jvm.java.is_empty", 0, 0),
        ("PriorityQueue", "add", "jvm.java.priority_add", 1, 1),
        ("PriorityQueue", "offer", "jvm.java.priority_add", 1, 1),
        ("PriorityQueue", "poll", "jvm.java.priority_poll", 0, 0),
        ("PriorityQueue", "remove", "jvm.java.priority_poll", 0, 0),
        ("PriorityQueue", "peek", "jvm.java.priority_peek", 0, 0),
    ];

    // Fold each type's ancestry, nearest first. `JAVA_TYPES` already carries
    // the chain — it is what stamps `__types` for `instanceof` — so the same
    // declaration drives identity and member resolution, and the two cannot
    // drift apart.
    //
    // `or_insert_with` is what makes nearest-declaration-win: the type's own
    // rows land before its interfaces', and a nearer interface before a
    // further one.
    // `util.concurrent` folds too: `CopyOnWriteArrayList` declares `List` in
    // its ancestry and `LinkedBlockingQueue` declares `BlockingQueue`/`Queue`,
    // so they inherit the same surface and an exact `== "util"` test silently
    // gave them none. Measured A/B on `java/copy_on_write_list`: 41/8 either
    // way — the concurrent types reach their own members through the Java
    // profile, so this neither shadows them nor is shadowed by them.
    for ty in JAVA_TYPES
        .iter()
        .filter(|t| t.package == "util" || t.package.starts_with("util."))
    {
        let mut methods = Subtree::new();
        for ancestor in ty.ancestry {
            for (owner, member, emit, min_args, max_args) in SPECS {
                if !owner.eq_ignore_ascii_case(ancestor) {
                    continue;
                }
                methods
                    .entry((*member).to_string())
                    .or_insert_with(|| common_method(emit, *min_args, *max_args));
            }
        }
        if methods.is_empty() {
            continue;
        }
        let type_path = format!("{}.{}", ty.package, ty.name.to_lowercase());
        ensure_type_node(root, &type_path);
        merge_type_methods(root, &type_path, methods);
    }

    merge_type_member_returns(
        root,
        "util.vector",
        &[
            ("iterator", "java.util.Iterator"),
            ("listIterator", "java.util.ListIterator"),
        ],
    );

    for (type_path, member, emit) in [
        ("util.hashmap", "keys", "jvm.java.map_key_set"),
        ("util.hashmap", "values", "jvm.java.map_values"),
        ("util.linkedhashmap", "keys", "jvm.java.map_key_set"),
        ("util.linkedhashmap", "values", "jvm.java.map_values"),
        ("util.treemap", "keys", "jvm.java.sorted_map_key_set"),
        ("util.treemap", "values", "jvm.java.sorted_map_values"),
    ] {
        let mut methods = Subtree::new();
        methods.insert(member.to_string(), common_emit(emit));
        merge_type_methods(root, type_path, methods);
    }
}

fn insert_java_util_regex(root: &mut Subtree) {
    ensure_type_node(root, "util.regex.pattern");
    ensure_type_node(root, "util.regex.matcher");

    insert_path(
        root,
        "util.regex.pattern.compile",
        common_method("jvm.java.regex_pattern_compile", 1, 2),
    );
    insert_path(
        root,
        "util.regex.pattern.quote",
        common_emit("strings.escape_regex"),
    );
    for (name, value) in [
        ("unix_lines", 1.0),
        ("case_insensitive", 2.0),
        ("comments", 4.0),
        ("multiline", 8.0),
        ("literal", 16.0),
        ("dotall", 32.0),
        ("unicode_case", 64.0),
        ("canonical_eq", 128.0),
        ("unicode_character_class", 256.0),
    ] {
        insert_path(
            root,
            &format!("util.regex.pattern.{name}"),
            NamespaceNode::Const(Value::F64(value)),
        );
    }
    merge_type_methods(
        root,
        "util.regex.pattern",
        [
            (
                "matcher".to_string(),
                common_method("jvm.java.regex_pattern_matcher", 1, 1),
            ),
            (
                "split".to_string(),
                common_method("jvm.java.regex_pattern_split", 1, 2),
            ),
            (
                "pattern".to_string(),
                common_method("jvm.java.regex_pattern_pattern", 0, 0),
            ),
            (
                "flags".to_string(),
                common_method("jvm.java.regex_pattern_flags", 0, 0),
            ),
            (
                "toString".to_string(),
                common_method("jvm.java.regex_pattern_pattern", 0, 0),
            ),
        ]
        .into_iter()
        .collect(),
    );
    merge_type_member_returns(
        root,
        "util.regex.pattern",
        &[
            ("matcher", "java.util.regex.Matcher"),
            ("split", "Array"),
            ("pattern", "String"),
            ("flags", "Int"),
            ("toString", "String"),
        ],
    );

    merge_type_methods(
        root,
        "util.regex.matcher",
        [
            (
                "find".to_string(),
                common_method("jvm.java.regex_matcher_find", 0, 0),
            ),
            (
                "matches".to_string(),
                common_method("jvm.java.regex_matcher_matches", 0, 0),
            ),
            (
                "lookingAt".to_string(),
                common_method("jvm.java.regex_matcher_find", 0, 0),
            ),
            (
                "group".to_string(),
                common_method("jvm.java.regex_matcher_group", 0, 1),
            ),
            (
                "start".to_string(),
                common_method("jvm.java.regex_matcher_start", 0, 0),
            ),
            (
                "end".to_string(),
                common_method("jvm.java.regex_matcher_end", 0, 0),
            ),
            (
                "reset".to_string(),
                common_method("jvm.java.regex_matcher_reset", 0, 1),
            ),
            (
                "replaceAll".to_string(),
                common_method("jvm.java.regex_pattern_replace_all", 1, 1),
            ),
            (
                "replaceFirst".to_string(),
                common_method("jvm.java.regex_pattern_replace_first", 1, 1),
            ),
            // `value`/`range` are PROPERTIES, not zero-arg methods. Kotlin
            // spells them as reads — `pattern.find(s)?.value` — and a member
            // READ never consults this `methods` subtree for a `Fn`/`CommonEmit`
            // leaf, so the read fell through to the dynamic path, produced the
            // matched text, and the surrounding force-call then CALLED it:
            // `string is not callable (type: 42)`.
            //
            // A `Property` leaf answers BOTH spellings: the shared member read
            // resolves it through `lookup_type_property_target`, and
            // `lookup_type_instance_target` unwraps `get` at argc 0, so Java's
            // own `m.value()` call form is unchanged.
            (
                "value".to_string(),
                namespaces::property(
                    Some(common_emit("jvm.java.regex_match_result_value")),
                    None,
                ),
            ),
            (
                "range".to_string(),
                namespaces::property(
                    Some(common_emit("jvm.java.regex_match_result_range")),
                    None,
                ),
            ),
        ]
        .into_iter()
        .collect(),
    );
    merge_type_member_returns(
        root,
        "util.regex.matcher",
        &[
            ("find", "Boolean"),
            ("matches", "Boolean"),
            ("lookingAt", "Boolean"),
            ("group", "String"),
            ("start", "Int"),
            ("end", "Int"),
            ("reset", "java.util.regex.Matcher"),
            ("replaceAll", "String"),
            ("replaceFirst", "String"),
            ("value", "String"),
            ("range", "IntRange"),
        ],
    );
}

fn insert_java_io_methods(root: &mut Subtree) {
    insert_common_static(
        root,
        "io.file",
        "createtempfile",
        "jvm.java.io_file_create_temp",
    );

    const SPECS: &[(&str, &str, &str, u8, u8)] = &[
        ("io.file", "writetext", "jvm.java.io_file_write_text", 1, 2),
        (
            "io.file",
            "appendtext",
            "jvm.java.io_file_append_text",
            1,
            2,
        ),
        ("io.file", "readtext", "jvm.java.io_file_read_text", 0, 1),
        ("io.file", "readlines", "jvm.java.io_file_read_lines", 0, 1),
        ("io.file", "listfiles", "jvm.java.io_file_list_files", 0, 1),
        ("io.file", "walk", "jvm.java.io_file_walk", 0, 0),
        ("io.file", "walktopdown", "jvm.java.io_file_walk", 0, 0),
        ("io.file", "walkbottomup", "jvm.java.io_file_walk", 0, 0),
        ("io.file", "writebytes", "jvm.java.io_file_write_text", 1, 1),
        ("io.file", "readbytes", "jvm.java.io_file_read_text", 0, 0),
        ("io.file", "exists", "jvm.java.io_file_exists", 0, 0),
        ("io.file", "delete", "jvm.java.io_file_delete", 0, 0),
        ("io.file", "deleteonexit", "jvm.java.io_false", 0, 0),
        ("io.file", "mkdir", "jvm.java.io_file_mkdirs", 0, 0),
        ("io.file", "mkdirs", "jvm.java.io_file_mkdirs", 0, 0),
        ("io.file", "isfile", "jvm.java.io_file_is_file", 0, 0),
        (
            "io.file",
            "isdirectory",
            "jvm.java.io_file_is_directory",
            0,
            0,
        ),
        ("io.file", "canread", "jvm.java.io_file_exists", 0, 0),
        ("io.file", "canwrite", "jvm.java.io_file_exists", 0, 0),
        ("io.file", "canexecute", "jvm.java.io_false", 0, 0),
        ("io.file", "getpath", "jvm.java.io_file_get_path", 0, 0),
        ("io.file", "path", "jvm.java.io_file_get_path", 0, 0),
        (
            "io.file",
            "getabsolutepath",
            "jvm.java.io_file_get_path",
            0,
            0,
        ),
        ("io.file", "absolutepath", "jvm.java.io_file_get_path", 0, 0),
        ("io.file", "getabsolutefile", "jvm.java.identity", 0, 0),
        ("io.file", "absolutefile", "jvm.java.identity", 0, 0),
        ("io.file", "getname", "jvm.java.io_file_get_name", 0, 0),
        ("io.file", "name", "jvm.java.io_file_get_name", 0, 0),
        ("io.file", "filename", "jvm.java.io_file_get_name", 0, 0),
        ("io.file", "extension", "jvm.java.io_file_extension", 0, 0),
        (
            "io.file",
            "namewithoutextension",
            "jvm.java.io_file_name_without_extension",
            0,
            0,
        ),
        ("io.file", "getparent", "jvm.java.io_file_parent", 0, 0),
        ("io.file", "parent", "jvm.java.io_file_parent", 0, 0),
        (
            "io.file",
            "getparentfile",
            "jvm.java.io_file_parent_file",
            0,
            0,
        ),
        (
            "io.file",
            "parentfile",
            "jvm.java.io_file_parent_file",
            0,
            0,
        ),
        ("io.file", "topath", "jvm.java.identity", 0, 0),
        ("io.file", "touri", "jvm.java.io_file_get_path", 0, 0),
        ("io.file", "tostring", "jvm.java.io_file_get_path", 0, 0),
        (
            "io.file",
            "lastmodified",
            "jvm.java.current_time_millis",
            0,
            0,
        ),
        ("io.file", "renameto", "jvm.java.io_file_rename_to", 1, 1),
        ("io.file", "copyto", "jvm.java.io_file_copy_to", 1, 2),
        (
            "io.file",
            "inputstream",
            "jvm.java.io_file_input_stream",
            0,
            0,
        ),
        (
            "io.file",
            "outputstream",
            "jvm.java.io_file_output_stream",
            0,
            0,
        ),
        (
            "io.file",
            "appendstream",
            "jvm.java.io_file_append_stream",
            0,
            0,
        ),
        ("io.file", "reader", "jvm.java.io_file_input_stream", 0, 0),
        ("io.file", "writer", "jvm.java.io_file_output_stream", 0, 0),
        ("io.bytearrayoutputstream", "size", "jvm.java.io_size", 0, 0),
        ("io.bytearrayoutputstream", "use", "jvm.java.io_use", 1, 1),
        (
            "io.bytearrayoutputstream",
            "flush",
            "jvm.java.io_flush_close",
            0,
            0,
        ),
        (
            "io.bytearrayoutputstream",
            "close",
            "jvm.java.io_flush_close",
            0,
            0,
        ),
        (
            "io.bytearrayoutputstream",
            "tostring",
            "jvm.java.io_output_to_string",
            0,
            1,
        ),
        (
            "io.bytearrayoutputstream",
            "write",
            "jvm.java.io_output_write",
            1,
            3,
        ),
        (
            "io.bytearrayoutputstream",
            "reset",
            "jvm.java.io_reset_buffer",
            0,
            0,
        ),
        (
            "io.bytearrayoutputstream",
            "tobytearray",
            "jvm.java.io_to_byte_array",
            0,
            0,
        ),
        ("io.bytearrayinputstream", "read", "jvm.java.io_read", 0, 1),
        (
            "io.bytearrayinputstream",
            "copyto",
            "jvm.java.io_stream_copy_to",
            1,
            1,
        ),
        ("io.bytearrayinputstream", "use", "jvm.java.io_use", 1, 1),
        (
            "io.bytearrayinputstream",
            "close",
            "jvm.java.io_flush_close",
            0,
            0,
        ),
        (
            "io.bytearrayinputstream",
            "available",
            "jvm.java.io_available",
            0,
            0,
        ),
        ("io.bytearrayinputstream", "mark", "jvm.java.io_mark", 1, 1),
        (
            "io.bytearrayinputstream",
            "reset",
            "jvm.java.io_reset_pos",
            0,
            0,
        ),
        (
            "io.bytearrayinputstream",
            "marksupported",
            "jvm.java.io_mark_supported",
            0,
            0,
        ),
        ("io.bytearrayinputstream", "skip", "jvm.java.io_skip", 1, 1),
        ("io.stringreader", "read", "jvm.java.io_read", 0, 1),
        ("io.stringreader", "use", "jvm.java.io_use", 1, 1),
        ("io.stringreader", "close", "jvm.java.io_flush_close", 0, 0),
        ("io.stringreader", "mark", "jvm.java.io_mark", 1, 1),
        ("io.stringreader", "reset", "jvm.java.io_reset_pos", 0, 0),
        ("io.stringreader", "skip", "jvm.java.io_skip", 1, 1),
        ("io.chararrayreader", "read", "jvm.java.io_read", 0, 1),
        ("io.chararrayreader", "use", "jvm.java.io_use", 1, 1),
        (
            "io.chararrayreader",
            "close",
            "jvm.java.io_flush_close",
            0,
            0,
        ),
        ("io.chararrayreader", "mark", "jvm.java.io_mark", 1, 1),
        ("io.chararrayreader", "reset", "jvm.java.io_reset_pos", 0, 0),
        ("io.chararrayreader", "skip", "jvm.java.io_skip", 1, 1),
        ("io.inputstreamreader", "read", "jvm.java.io_read", 0, 1),
        (
            "io.inputstreamreader",
            "readtext",
            "jvm.java.io_read_text",
            0,
            0,
        ),
        ("io.inputstreamreader", "use", "jvm.java.io_use", 1, 1),
        (
            "io.inputstreamreader",
            "close",
            "jvm.java.io_flush_close",
            0,
            0,
        ),
        ("io.inputstreamreader", "ready", "jvm.java.io_ready", 0, 0),
        ("io.bufferedinputstream", "read", "jvm.java.io_read", 0, 1),
        (
            "io.bufferedinputstream",
            "readtext",
            "jvm.java.io_read_text",
            0,
            0,
        ),
        (
            "io.bufferedinputstream",
            "copyto",
            "jvm.java.io_stream_copy_to",
            1,
            1,
        ),
        ("io.bufferedinputstream", "use", "jvm.java.io_use", 1, 1),
        (
            "io.bufferedinputstream",
            "close",
            "jvm.java.io_flush_close",
            0,
            0,
        ),
        ("io.filterinputstream", "read", "jvm.java.io_read", 0, 1),
        (
            "io.filterinputstream",
            "readtext",
            "jvm.java.io_read_text",
            0,
            0,
        ),
        (
            "io.filterinputstream",
            "copyto",
            "jvm.java.io_stream_copy_to",
            1,
            1,
        ),
        ("io.filterinputstream", "use", "jvm.java.io_use", 1, 1),
        (
            "io.filterinputstream",
            "close",
            "jvm.java.io_flush_close",
            0,
            0,
        ),
        ("io.pushbackinputstream", "read", "jvm.java.io_read", 0, 0),
        (
            "io.pushbackinputstream",
            "unread",
            "jvm.java.io_unread",
            1,
            1,
        ),
        ("io.bufferedreader", "read", "jvm.java.io_read", 0, 1),
        (
            "io.bufferedreader",
            "readtext",
            "jvm.java.io_read_text",
            0,
            0,
        ),
        ("io.bufferedreader", "use", "jvm.java.io_use", 1, 1),
        (
            "io.bufferedreader",
            "close",
            "jvm.java.io_flush_close",
            0,
            0,
        ),
        (
            "io.bufferedreader",
            "readline",
            "jvm.java.io_read_line",
            0,
            0,
        ),
        ("io.bufferedreader", "ready", "jvm.java.io_ready", 0, 0),
        (
            "io.bufferedreader",
            "marksupported",
            "jvm.java.io_mark_supported",
            0,
            0,
        ),
        (
            "io.linenumberreader",
            "readline",
            "jvm.java.io_read_line",
            0,
            0,
        ),
        (
            "io.linenumberreader",
            "getlinenumber",
            "jvm.java.io_get_line_number",
            0,
            0,
        ),
        ("io.printwriter", "print", "jvm.java.io_writer_print", 1, 1),
        (
            "io.printwriter",
            "println",
            "jvm.java.io_writer_println",
            0,
            1,
        ),
        (
            "io.printwriter",
            "append",
            "jvm.java.io_writer_append",
            1,
            1,
        ),
        ("io.printwriter", "flush", "jvm.java.io_flush_close", 0, 0),
        ("io.printwriter", "close", "jvm.java.io_flush_close", 0, 0),
        ("io.printwriter", "checkerror", "jvm.java.io_false", 0, 0),
        ("io.printstream", "print", "jvm.java.io_writer_print", 1, 1),
        (
            "io.printstream",
            "println",
            "jvm.java.io_writer_println",
            0,
            1,
        ),
        ("io.printstream", "flush", "jvm.java.io_flush_close", 0, 0),
        (
            "io.outputstreamwriter",
            "write",
            "jvm.java.io_writer_write",
            1,
            3,
        ),
        ("io.outputstreamwriter", "use", "jvm.java.io_use", 1, 1),
        (
            "io.outputstreamwriter",
            "flush",
            "jvm.java.io_flush_close",
            0,
            0,
        ),
        (
            "io.outputstreamwriter",
            "close",
            "jvm.java.io_flush_close",
            0,
            0,
        ),
        ("io.stringwriter", "write", "jvm.java.io_writer_write", 1, 3),
        ("io.stringwriter", "use", "jvm.java.io_use", 1, 1),
        (
            "io.stringwriter",
            "append",
            "jvm.java.io_writer_append",
            1,
            1,
        ),
        (
            "io.stringwriter",
            "tostring",
            "jvm.java.io_writer_to_string",
            0,
            0,
        ),
        ("io.stringwriter", "flush", "jvm.java.io_flush_close", 0, 0),
        ("io.stringwriter", "close", "jvm.java.io_flush_close", 0, 0),
        (
            "io.chararraywriter",
            "write",
            "jvm.java.io_writer_write",
            1,
            3,
        ),
        ("io.chararraywriter", "use", "jvm.java.io_use", 1, 1),
        (
            "io.chararraywriter",
            "tochararray",
            "jvm.java.io_writer_to_char_array",
            0,
            0,
        ),
        (
            "io.chararraywriter",
            "tostring",
            "jvm.java.io_writer_to_string",
            0,
            0,
        ),
        (
            "io.chararraywriter",
            "flush",
            "jvm.java.io_flush_close",
            0,
            0,
        ),
        (
            "io.chararraywriter",
            "close",
            "jvm.java.io_flush_close",
            0,
            0,
        ),
        (
            "io.bufferedwriter",
            "write",
            "jvm.java.io_writer_write",
            1,
            3,
        ),
        (
            "io.bufferedwriter",
            "newline",
            "jvm.java.io_writer_newline",
            0,
            0,
        ),
        (
            "io.bufferedwriter",
            "flush",
            "jvm.java.io_flush_close",
            0,
            0,
        ),
        ("io.filterwriter", "write", "jvm.java.io_writer_write", 1, 3),
        (
            "io.filterwriter",
            "write$sigint",
            "jvm.java.io_writer_write",
            1,
            1,
        ),
        (
            "io.filterwriter",
            "write$sigchararray_int_int",
            "jvm.java.io_writer_write",
            3,
            3,
        ),
        (
            "io.filterwriter",
            "write$sigstring_int_int",
            "jvm.java.io_writer_write",
            3,
            3,
        ),
        ("io.filterwriter", "flush", "jvm.java.io_flush_close", 0, 0),
        ("io.filterwriter", "close", "jvm.java.io_flush_close", 0, 0),
        (
            "io.dataoutputstream",
            "writeint",
            "jvm.java.io_output_write",
            1,
            1,
        ),
        (
            "io.dataoutputstream",
            "writelong",
            "jvm.java.io_output_write",
            1,
            1,
        ),
        (
            "io.dataoutputstream",
            "writeboolean",
            "jvm.java.io_output_write",
            1,
            1,
        ),
        (
            "io.dataoutputstream",
            "writeutf",
            "jvm.java.io_output_write",
            1,
            1,
        ),
        (
            "io.dataoutputstream",
            "flush",
            "jvm.java.io_flush_close",
            0,
            0,
        ),
        ("io.datainputstream", "readint", "jvm.java.io_read", 0, 0),
        ("io.datainputstream", "readlong", "jvm.java.io_read", 0, 0),
        (
            "io.datainputstream",
            "readboolean",
            "jvm.java.io_read",
            0,
            0,
        ),
        (
            "io.datainputstream",
            "readutf",
            "jvm.java.io_read_utf",
            0,
            0,
        ),
        ("io.sequenceinputstream", "read", "jvm.java.io_read", 0, 0),
    ];

    let mut by_type: BTreeMap<&str, Subtree> = BTreeMap::new();
    for (type_path, member, emit, min_args, max_args) in SPECS {
        by_type.entry(*type_path).or_default().insert(
            (*member).to_string(),
            common_method(emit, *min_args, *max_args),
        );
    }
    for (type_path, methods) in by_type {
        ensure_type_node(root, type_path);
        merge_type_methods(root, type_path, methods);
    }

    for (type_path, returns) in [
        ("io.printwriter", &[("append", "java.io.PrintWriter")][..]),
        ("io.stringwriter", &[("append", "java.io.StringWriter")][..]),
        (
            "io.file",
            &[
                ("createtempfile", "java.io.File"),
                ("copyto", "java.io.File"),
                ("inputstream", "java.io.ByteArrayInputStream"),
                ("outputstream", "java.io.ByteArrayOutputStream"),
                ("appendstream", "java.io.ByteArrayOutputStream"),
                ("reader", "java.io.InputStreamReader"),
                ("writer", "java.io.OutputStreamWriter"),
                ("getparentfile", "java.io.File"),
                ("parentfile", "java.io.File"),
                ("getabsolutefile", "java.io.File"),
                ("absolutefile", "java.io.File"),
                ("topath", "java.io.File"),
            ][..],
        ),
    ] {
        merge_type_member_returns(root, type_path, returns);
    }
    merge_type_member_returns(
        root,
        "io.file",
        &[
            ("exists", "Boolean"),
            ("delete", "Boolean"),
            ("mkdir", "Boolean"),
            ("mkdirs", "Boolean"),
            ("isfile", "Boolean"),
            ("isdirectory", "Boolean"),
            ("canread", "Boolean"),
            ("canwrite", "Boolean"),
            ("canexecute", "Boolean"),
        ],
    );
}

/// `java.util.EnumSet` and `java.lang.Enum` — declared as TREE data, children
/// and all.
///
/// This is what "migrated" means for a JDK class, and it is the difference
/// between a platform surface and a code move. Fifteen `common:` emits sitting
/// in `platforms/jvm` were still reached only through fifteen
/// `__java_enum_set_*` rows in the Java profile plus a Java-walker rewrite, so
/// `java.util.EnumSet.of(Color.RED)` from Kotlin resolved to nothing —
/// measured with `--dump-ast`: the member chain stayed raw.
///
/// A type node alone would not fix that. The statics have to live under
/// `statics`, the instance operations under `methods`, and the statics have to
/// DECLARE that they return a `java.util.EnumSet` (`member_returns`) so
/// `EnumSet.of(x).add(y)` — and a local typed by that call — resolves the
/// second hop through the tree as well.
fn insert_java_util_enum_set(root: &mut Subtree) {
    const STATICS: &[(&str, &str, u8, u8)] = &[
        ("noneof", "jvm.java.enum_set_none_of", 1, 1),
        ("allof", "jvm.java.enum_set_all_of", 1, 1),
        ("of", "jvm.java.enum_set_of", 1, 10),
        ("copyof", "jvm.java.enum_set_copy_of", 1, 1),
        ("complementof", "jvm.java.enum_set_complement_of", 1, 1),
        ("range", "jvm.java.enum_set_range", 2, 2),
    ];
    const METHODS: &[(&str, &str, u8, u8)] = &[
        ("add", "jvm.java.enum_set_add", 1, 1),
        ("addall", "jvm.java.enum_set_add_all", 1, 1),
        ("contains", "jvm.java.enum_set_contains", 1, 1),
        ("containsall", "jvm.java.enum_set_contains_all", 1, 1),
        ("remove", "jvm.java.enum_set_remove", 1, 1),
        ("equals", "jvm.java.enum_set_equals", 1, 1),
        ("hashcode", "jvm.java.enum_set_hash_code", 0, 0),
        ("iterator", "jvm.java.enum_set_iterator", 0, 0),
    ];

    for (member, emit, min_args, max_args) in STATICS {
        insert_path(
            root,
            &format!("util.enumset.{member}"),
            common_method(emit, *min_args, *max_args),
        );
    }
    // AFTER the statics, never before: `ensure_type_node` PROMOTES an existing
    // namespace and silently does nothing when the path is absent. Called
    // first it left `enumset` a plain `Namespace`, so the statics still
    // resolved by path walk while `find_type_node` — which matches `Type`
    // nodes only — saw no type at all, and every instance method and declared
    // return type was dropped without a word.
    ensure_type_node(root, "util.enumset");
    let mut methods = Subtree::new();
    for (member, emit, min_args, max_args) in METHODS {
        methods.insert(
            (*member).to_string(),
            common_method(emit, *min_args, *max_args),
        );
    }
    merge_type_methods(root, "util.enumset", methods);
    merge_type_member_returns(
        root,
        "util.enumset",
        &[
            ("noneOf", "java.util.EnumSet"),
            ("allOf", "java.util.EnumSet"),
            ("of", "java.util.EnumSet"),
            ("copyOf", "java.util.EnumSet"),
            ("complementOf", "java.util.EnumSet"),
            ("range", "java.util.EnumSet"),
            ("iterator", "java.util.Iterator"),
        ],
    );

    // `Class.getEnumConstants()`, and the private hook each enum's static
    // initializer calls to publish its own. Both live on `java.lang.Enum`
    // because that is the class the metadata belongs to — and because
    // reaching them by their real JDK path is exactly what makes them
    // available to a frontend that declares nothing.
    insert_path(
        root,
        "lang.enum.__vybe_declare",
        common_method("jvm.java.enum_declare", 2, 2),
    );
    insert_path(
        root,
        "lang.enum.getenumconstants",
        common_method("jvm.java.enum_constants_of", 1, 1),
    );
    ensure_type_node(root, "lang.enum");
}

fn insert_java_util_collection_statics(root: &mut Subtree) {
    for (type_path, member, emit) in [
        ("util.collections", "sort", "jvm.java.collections_sort"),
        (
            "util.collections",
            "reverse",
            "jvm.java.collections_reverse",
        ),
        (
            "util.collections",
            "shuffle",
            "jvm.java.collections_shuffle",
        ),
        ("util.collections", "fill", "jvm.java.collections_fill"),
        ("util.collections", "copy", "jvm.java.collections_copy"),
        ("util.collections", "addall", "jvm.java.collections_add_all"),
        ("util.collections", "rotate", "jvm.java.collections_rotate"),
        (
            "util.collections",
            "replaceall",
            "jvm.java.collections_replace_all",
        ),
        ("util.collections", "swap", "jvm.java.collections_swap"),
        (
            "util.collections",
            "indexofsublist",
            "jvm.java.collections_index_of_sublist",
        ),
        (
            "util.collections",
            "lastindexofsublist",
            "jvm.java.collections_last_index_of_sublist",
        ),
        ("util.collections", "min", "jvm.java.collections_min"),
        ("util.collections", "max", "jvm.java.collections_max"),
        (
            "util.collections",
            "frequency",
            "jvm.java.collections_frequency",
        ),
        (
            "util.collections",
            "disjoint",
            "jvm.java.collections_disjoint",
        ),
        (
            "util.collections",
            "reverseorder",
            "jvm.java.collections_reverse_order",
        ),
        (
            "util.collections",
            "newsetfrommap",
            "jvm.java.new_set_from_map",
        ),
        (
            "util.collections",
            "binarysearch",
            "jvm.java.arrays_binary_search",
        ),
        (
            "util.collections",
            "unmodifiablelist",
            "jvm.java.unmodifiable_list",
        ),
        (
            "util.collections",
            "unmodifiableset",
            "jvm.java.unmodifiable_set",
        ),
        (
            "util.collections",
            "unmodifiablemap",
            "jvm.java.unmodifiable_map",
        ),
        ("util.collections", "synchronizedlist", "jvm.java.identity"),
        ("util.collections", "singleton", "jvm.java.singleton_list"),
        ("util.collections", "singletonmap", "jvm.java.map_of"),
        (
            "util.collections",
            "singletonlist",
            "jvm.java.singleton_list",
        ),
        ("util.collections", "emptylist", "jvm.java.empty_list"),
        ("util.collections", "emptyset", "jvm.java.empty_set"),
        ("util.collections", "emptymap", "jvm.java.map_of"),
        ("util.collections", "ncopies", "jvm.java.n_copies"),
        ("util.arrays", "sort", "jvm.java.arrays_sort"),
        ("util.arrays", "parallelsort", "jvm.java.arrays_sort"),
        ("util.arrays", "fill", "jvm.java.arrays_fill"),
        ("util.arrays", "copyof", "jvm.java.arrays_copy_of"),
        (
            "util.arrays",
            "copyofrange",
            "jvm.java.arrays_copy_of_range",
        ),
        ("util.arrays", "tostring", "jvm.java.arrays_to_string"),
        (
            "util.arrays",
            "deeptostring",
            "jvm.java.arrays_deep_to_string",
        ),
        ("util.arrays", "aslist", "jvm.java.arrays_as_list"),
        ("util.arrays", "equals", "jvm.java.arrays_equals"),
        ("util.arrays", "deepequals", "jvm.java.arrays_deep_equals"),
        ("util.arrays", "compare", "jvm.java.arrays_compare"),
        (
            "util.arrays",
            "compareunsigned",
            "jvm.java.arrays_compare_unsigned",
        ),
        ("util.arrays", "mismatch", "jvm.java.arrays_mismatch"),
        ("util.arrays", "setall", "jvm.java.arrays_set_all"),
        ("util.arrays", "parallelsetall", "jvm.java.arrays_set_all"),
        (
            "util.arrays",
            "parallelprefix",
            "jvm.java.arrays_parallel_prefix",
        ),
        ("util.arrays", "hashcode", "jvm.java.arrays_hash_code"),
        (
            "util.arrays",
            "deephashcode",
            "jvm.java.arrays_deep_hash_code",
        ),
        (
            "util.arrays",
            "binarysearch",
            "jvm.java.arrays_binary_search",
        ),
        ("util.arrays", "stream", "jvm.java.identity"),
        ("util.list", "of", "jvm.java.list_of"),
        ("util.list", "copyof", "jvm.java.list_copy_of"),
        ("util.set", "of", "jvm.java.set_of"),
        ("util.set", "copyof", "jvm.java.set_copy_of"),
        ("util.map", "of", "jvm.java.map_of"),
        ("util.map", "entry", "jvm.java.map_entry"),
        ("util.map", "ofentries", "jvm.java.map_of_entries"),
        ("lang.string", "format", "sprintf.format"),
        ("lang.string", "join", "jvm.java.string_join"),
        (
            "util.objects",
            "requirenonnull",
            "jvm.java.require_non_null",
        ),
        ("util.optional", "empty", "jvm.java.optional_empty"),
        ("util.optional", "of", "jvm.java.optional_of"),
        (
            "util.optional",
            "ofnullable",
            "jvm.java.optional_of_nullable",
        ),
    ] {
        insert_common_static(root, type_path, member, emit);
    }
}

fn insert_java_time_statics(root: &mut Subtree) {
    for (type_path, member, emit) in [
        (
            "time.instant",
            "ofepochsecond",
            "jvm.java.instant_of_epoch_second",
        ),
        (
            "time.instant",
            "ofepochmilli",
            "jvm.java.instant_of_epoch_milli",
        ),
        ("time.instant", "parse", "jvm.java.instant_parse"),
        ("time.localdate", "of", "jvm.java.local_date_of"),
        ("time.localdate", "parse", "jvm.java.local_date_parse"),
        ("time.localtime", "of", "jvm.java.local_time_of"),
        ("time.localtime", "parse", "jvm.java.local_time_parse"),
        ("time.localdatetime", "of", "jvm.java.local_datetime_of"),
        (
            "time.localdatetime",
            "parse",
            "jvm.java.local_datetime_parse",
        ),
        ("time.offsetdatetime", "of", "jvm.java.offset_datetime_of"),
        (
            "time.offsetdatetime",
            "ofinstant",
            "jvm.java.offset_datetime_of_instant",
        ),
        (
            "time.offsetdatetime",
            "parse",
            "jvm.java.offset_datetime_parse",
        ),
        ("time.zoneddatetime", "of", "jvm.java.zoned_datetime_of"),
        (
            "time.zoneddatetime",
            "ofinstant",
            "jvm.java.zoned_datetime_of_instant",
        ),
        (
            "time.zoneddatetime",
            "ofstrict",
            "jvm.java.zoned_datetime_of_strict",
        ),
        (
            "time.zoneddatetime",
            "parse",
            "jvm.java.zoned_datetime_parse",
        ),
        ("time.period", "ofdays", "jvm.java.period_of_days"),
        ("time.period", "ofmonths", "jvm.java.period_of_months"),
        ("time.period", "between", "jvm.java.period_between"),
        ("time.duration", "ofhours", "jvm.java.duration_of_hours"),
        ("time.duration", "ofminutes", "jvm.java.duration_of_minutes"),
        ("time.duration", "ofseconds", "jvm.java.duration_of_seconds"),
        ("time.duration", "between", "jvm.java.duration_between"),
        ("time.yearmonth", "parse", "jvm.java.year_month_parse"),
        ("time.monthday", "parse", "jvm.java.month_day_parse"),
        (
            "time.zoneoffset",
            "ofhours",
            "jvm.java.zone_offset_of_hours",
        ),
        ("time.zoneid", "of", "jvm.java.zone_id_of"),
        (
            "time.zoneid",
            "systemdefault",
            "jvm.java.zone_id_system_default",
        ),
        ("time.zoneid", "from", "jvm.java.zone_id_from"),
        ("time.zoneid", "ofoffset", "jvm.java.zone_id_of_offset"),
        ("time.instant", "now", "jvm.java.instant_now"),
        ("time.clock", "fixed", "jvm.java.clock_fixed"),
        ("time.clock", "systemutc", "jvm.java.identity"),
        ("time.duration", "parse", "jvm.java.duration_parse"),
        (
            "time.temporal.chronounit.days",
            "between",
            "jvm.java.chrono_days_between",
        ),
        (
            "time.temporal.chronounit.weeks",
            "between",
            "jvm.java.chrono_weeks_between",
        ),
        (
            "time.temporal.chronounit.months",
            "between",
            "jvm.java.chrono_months_between",
        ),
    ] {
        insert_common_static(root, type_path, member, emit);
    }

    for (type_path, returns) in [
        (
            "time.instant",
            &[
                ("ofEpochSecond", "java.time.Instant"),
                ("ofEpochMilli", "java.time.Instant"),
                ("parse", "java.time.Instant"),
                ("now", "java.time.Instant"),
            ][..],
        ),
        ("time.clock", &[("fixed", "java.time.Clock")][..]),
        (
            "time.localdate",
            &[
                ("of", "java.time.LocalDate"),
                ("parse", "java.time.LocalDate"),
            ][..],
        ),
        (
            "time.localtime",
            &[
                ("of", "java.time.LocalTime"),
                ("parse", "java.time.LocalTime"),
            ][..],
        ),
        (
            "time.localdatetime",
            &[
                ("of", "java.time.LocalDateTime"),
                ("parse", "java.time.LocalDateTime"),
            ][..],
        ),
        (
            "time.duration",
            &[
                ("ofHours", "java.time.Duration"),
                ("ofMinutes", "java.time.Duration"),
                ("ofSeconds", "java.time.Duration"),
                ("parse", "java.time.Duration"),
                ("between", "java.time.Duration"),
            ][..],
        ),
        ("time.yearmonth", &[("parse", "java.time.YearMonth")][..]),
        ("time.monthday", &[("parse", "java.time.MonthDay")][..]),
        (
            "time.period",
            &[
                ("ofDays", "java.time.Period"),
                ("ofMonths", "java.time.Period"),
                ("between", "java.time.Period"),
            ][..],
        ),
        (
            "time.zoneid",
            &[
                ("of", "java.time.ZoneId"),
                ("systemDefault", "java.time.ZoneId"),
                ("from", "java.time.ZoneId"),
                ("ofOffset", "java.time.ZoneId"),
            ][..],
        ),
        (
            "time.zoneoffset",
            &[("ofHours", "java.time.ZoneOffset")][..],
        ),
        (
            "time.offsetdatetime",
            &[
                ("of", "java.time.OffsetDateTime"),
                ("ofInstant", "java.time.OffsetDateTime"),
                ("parse", "java.time.OffsetDateTime"),
            ][..],
        ),
        (
            "time.zoneddatetime",
            &[
                ("of", "java.time.ZonedDateTime"),
                ("ofInstant", "java.time.ZonedDateTime"),
                ("ofStrict", "java.time.ZonedDateTime"),
                ("parse", "java.time.ZonedDateTime"),
            ][..],
        ),
    ] {
        merge_type_member_returns(root, type_path, returns);
    }
}

fn insert_java_time_instance_members(root: &mut Subtree) {
    let prop = |emit: &str| namespaces::property(Some(common_emit(emit)), None);
    let mut date = Subtree::new();
    for (name, emit) in [
        ("year", "jvm.java.instant_get_year"),
        ("monthvalue", "jvm.java.instant_get_month"),
        ("dayofmonth", "jvm.java.instant_get_day"),
        ("dayofyear", "jvm.java.time_day_of_year"),
        ("dayofweek", "jvm.java.time_day_of_week"),
    ] {
        date.insert(name.to_string(), prop(emit));
    }
    for (name, emit, min_args, max_args) in [
        ("getyear", "jvm.java.instant_get_year", 0, 0),
        ("getmonthvalue", "jvm.java.instant_get_month", 0, 0),
        ("getdayofmonth", "jvm.java.instant_get_day", 0, 0),
        ("getdayofyear", "jvm.java.time_day_of_year", 0, 0),
        ("getdayofweek", "jvm.java.time_day_of_week", 0, 0),
        ("get", "jvm.java.time_get", 1, 1),
        ("isleapyear", "jvm.java.time_is_leap_year", 0, 0),
        ("plusdays", "jvm.java.time_plus_days", 1, 1),
        ("minusdays", "jvm.java.time_minus_days", 1, 1),
        ("plusweeks", "jvm.java.time_plus_weeks", 1, 1),
        ("plusmonths", "jvm.java.time_plus_months", 1, 1),
        ("minusmonths", "jvm.java.time_minus_months", 1, 1),
        ("withyear", "jvm.java.time_with_year", 1, 1),
        ("withmonth", "jvm.java.time_with_month", 1, 1),
        ("withdayofmonth", "jvm.java.time_with_day", 1, 1),
        ("lengthofmonth", "jvm.java.time_length_of_month", 0, 0),
        ("isbefore", "jvm.java.instant_is_before", 1, 1),
        ("isafter", "jvm.java.instant_is_after", 1, 1),
        ("compareto", "jvm.java.instant_compare_to", 1, 1),
        ("tostring", "jvm.java.time_to_string", 0, 0),
    ] {
        date.insert(name.to_string(), common_method(emit, min_args, max_args));
    }
    ensure_type_node(root, "time.localdate");
    merge_type_methods(root, "time.localdate", date);
    merge_type_member_returns(
        root,
        "time.localdate",
        &[
            ("dayOfWeek", "java.time.DayOfWeek"),
            ("getDayOfWeek", "java.time.DayOfWeek"),
            ("plusDays", "java.time.LocalDate"),
            ("minusDays", "java.time.LocalDate"),
            ("plusWeeks", "java.time.LocalDate"),
            ("plusMonths", "java.time.LocalDate"),
            ("minusMonths", "java.time.LocalDate"),
            ("withYear", "java.time.LocalDate"),
            ("withMonth", "java.time.LocalDate"),
            ("withDayOfMonth", "java.time.LocalDate"),
        ],
    );

    let mut time = Subtree::new();
    for (name, emit) in [
        ("hour", "jvm.java.instant_get_hour"),
        ("minute", "jvm.java.instant_get_minute"),
        ("second", "jvm.java.instant_get_second"),
    ] {
        time.insert(name.to_string(), prop(emit));
    }
    for (name, emit, min_args, max_args) in [
        ("gethour", "jvm.java.instant_get_hour", 0, 0),
        ("getminute", "jvm.java.instant_get_minute", 0, 0),
        ("getsecond", "jvm.java.instant_get_second", 0, 0),
        ("plushours", "jvm.java.time_plus_hours", 1, 1),
        ("minushours", "jvm.java.time_minus_hours", 1, 1),
        ("plusminutes", "jvm.java.time_plus_minutes", 1, 1),
        ("minusminutes", "jvm.java.time_minus_minutes", 1, 1),
        ("plusseconds", "jvm.java.time_plus_seconds", 1, 1),
        ("minusseconds", "jvm.java.time_minus_seconds", 1, 1),
        ("withhour", "jvm.java.time_with_hour", 1, 1),
        ("withminute", "jvm.java.time_with_minute", 1, 1),
        ("withsecond", "jvm.java.time_with_second", 1, 1),
        ("isbefore", "jvm.java.instant_is_before", 1, 1),
        ("isafter", "jvm.java.instant_is_after", 1, 1),
        ("compareto", "jvm.java.instant_compare_to", 1, 1),
        ("tostring", "jvm.java.timeofday_to_string", 0, 0),
    ] {
        time.insert(name.to_string(), common_method(emit, min_args, max_args));
    }
    ensure_type_node(root, "time.localtime");
    merge_type_methods(root, "time.localtime", time);
    merge_type_member_returns(
        root,
        "time.localtime",
        &[
            ("plusHours", "java.time.LocalTime"),
            ("minusHours", "java.time.LocalTime"),
            ("plusMinutes", "java.time.LocalTime"),
            ("minusMinutes", "java.time.LocalTime"),
            ("plusSeconds", "java.time.LocalTime"),
            ("minusSeconds", "java.time.LocalTime"),
            ("withHour", "java.time.LocalTime"),
            ("withMinute", "java.time.LocalTime"),
            ("withSecond", "java.time.LocalTime"),
        ],
    );

    let mut duration = Subtree::new();
    duration.insert("seconds".to_string(), prop("jvm.java.identity"));
    for (name, emit, min_args, max_args) in [
        ("getseconds", "jvm.java.identity", 0, 0),
        ("tominutes", "jvm.java.duration_to_minutes", 0, 0),
        ("tomillis", "jvm.java.duration_to_millis", 0, 0),
        ("tohours", "jvm.java.duration_to_hours", 0, 0),
        ("plushours", "jvm.java.duration_plus_hours", 1, 1),
        ("minushours", "jvm.java.duration_minus_hours", 1, 1),
        ("plusminutes", "jvm.java.duration_plus_minutes", 1, 1),
        ("minusminutes", "jvm.java.duration_minus_minutes", 1, 1),
        ("multipliedby", "jvm.java.duration_multiplied_by", 1, 1),
        ("negated", "jvm.java.duration_negated", 0, 0),
        ("iszero", "jvm.java.duration_is_zero", 0, 0),
        ("tostring", "jvm.java.time_to_string", 0, 0),
    ] {
        duration.insert(name.to_string(), common_method(emit, min_args, max_args));
    }
    ensure_type_node(root, "time.duration");
    merge_type_methods(root, "time.duration", duration);
    merge_type_member_returns(
        root,
        "time.duration",
        &[
            ("plusHours", "java.time.Duration"),
            ("minusHours", "java.time.Duration"),
            ("plusMinutes", "java.time.Duration"),
            ("minusMinutes", "java.time.Duration"),
            ("multipliedBy", "java.time.Duration"),
            ("negated", "java.time.Duration"),
        ],
    );

    let mut instant = Subtree::new();
    for (name, emit, min_args, max_args) in [
        ("getepochsecond", "jvm.java.instant_get_epoch_second", 0, 0),
        ("getnano", "jvm.java.instant_get_nano", 0, 0),
        ("toepochmilli", "jvm.java.instant_to_epoch_milli", 0, 0),
        ("plusseconds", "jvm.java.instant_plus_seconds", 1, 1),
        ("minusseconds", "jvm.java.instant_minus_seconds", 1, 1),
        ("plusmillis", "jvm.java.instant_plus_millis", 1, 1),
        ("minusmillis", "jvm.java.instant_minus_millis", 1, 1),
        ("plusnanos", "jvm.java.instant_plus_nanos", 1, 1),
        ("minusnanos", "jvm.java.instant_minus_nanos", 1, 1),
        ("isbefore", "jvm.java.instant_is_before", 1, 1),
        ("isafter", "jvm.java.instant_is_after", 1, 1),
        ("compareto", "jvm.java.instant_compare_to", 1, 1),
        ("tostring", "jvm.java.instant_to_string", 0, 0),
        ("toinstant", "jvm.java.identity", 0, 0),
    ] {
        instant.insert(name.to_string(), common_method(emit, min_args, max_args));
    }
    instant.insert(
        "epochsecond".to_string(),
        prop("jvm.java.instant_get_epoch_second"),
    );
    instant.insert("nano".to_string(), prop("jvm.java.instant_get_nano"));
    ensure_type_node(root, "time.instant");
    merge_type_methods(root, "time.instant", instant);
    merge_type_member_returns(
        root,
        "time.instant",
        &[
            ("plusSeconds", "java.time.Instant"),
            ("minusSeconds", "java.time.Instant"),
            ("plusMillis", "java.time.Instant"),
            ("minusMillis", "java.time.Instant"),
            ("plusNanos", "java.time.Instant"),
            ("minusNanos", "java.time.Instant"),
            ("toInstant", "java.time.Instant"),
        ],
    );

    let mut date_time = Subtree::new();
    for (name, emit) in [
        ("year", "jvm.java.instant_get_year"),
        ("monthvalue", "jvm.java.instant_get_month"),
        ("dayofmonth", "jvm.java.instant_get_day"),
        ("hour", "jvm.java.instant_get_hour"),
        ("minute", "jvm.java.instant_get_minute"),
        ("second", "jvm.java.instant_get_second"),
        ("date", "jvm.java.instant_to_local_date"),
        ("offset", "jvm.java.instant_get_offset"),
        ("zone", "jvm.java.instant_get_zone"),
    ] {
        date_time.insert(name.to_string(), prop(emit));
    }
    for (name, emit, min_args, max_args) in [
        ("getyear", "jvm.java.instant_get_year", 0, 0),
        ("getmonthvalue", "jvm.java.instant_get_month", 0, 0),
        ("getdayofmonth", "jvm.java.instant_get_day", 0, 0),
        ("gethour", "jvm.java.instant_get_hour", 0, 0),
        ("getminute", "jvm.java.instant_get_minute", 0, 0),
        ("getsecond", "jvm.java.instant_get_second", 0, 0),
        ("plusdays", "jvm.java.time_plus_days", 1, 1),
        ("minusdays", "jvm.java.time_minus_days", 1, 1),
        ("plusweeks", "jvm.java.time_plus_weeks", 1, 1),
        ("plusmonths", "jvm.java.time_plus_months", 1, 1),
        ("minusmonths", "jvm.java.time_minus_months", 1, 1),
        ("plushours", "jvm.java.time_plus_hours", 1, 1),
        ("minushours", "jvm.java.time_minus_hours", 1, 1),
        ("plusminutes", "jvm.java.time_plus_minutes", 1, 1),
        ("minusminutes", "jvm.java.time_minus_minutes", 1, 1),
        ("plusseconds", "jvm.java.time_plus_seconds", 1, 1),
        ("minusseconds", "jvm.java.time_minus_seconds", 1, 1),
        ("withyear", "jvm.java.time_with_year", 1, 1),
        ("withmonth", "jvm.java.time_with_month", 1, 1),
        ("withdayofmonth", "jvm.java.time_with_day", 1, 1),
        ("withhour", "jvm.java.time_with_hour", 1, 1),
        ("withminute", "jvm.java.time_with_minute", 1, 1),
        ("withsecond", "jvm.java.time_with_second", 1, 1),
        ("with", "jvm.java.time_with", 2, 2),
        ("get", "jvm.java.time_get", 1, 1),
        ("isbefore", "jvm.java.instant_is_before", 1, 1),
        ("isafter", "jvm.java.instant_is_after", 1, 1),
        ("compareto", "jvm.java.instant_compare_to", 1, 1),
        ("tostring", "jvm.java.datetime_to_string", 0, 0),
        ("tolocaldate", "jvm.java.instant_to_local_date", 0, 0),
        ("tolocaltime", "jvm.java.identity", 0, 0),
        ("tolocaldatetime", "jvm.java.identity", 0, 0),
        ("toinstant", "jvm.java.identity", 0, 0),
        ("tooffsetdatetime", "jvm.java.identity", 0, 0),
        ("tozoneddatetime", "jvm.java.identity", 0, 0),
        ("atzonesameinstant", "jvm.java.instant_with_offset", 1, 1),
        (
            "atzonesimilarlocal",
            "jvm.java.time_with_offset_same_local",
            1,
            1,
        ),
        (
            "withoffsetsameinstant",
            "jvm.java.instant_with_offset",
            1,
            1,
        ),
        (
            "withoffsetsamelocal",
            "jvm.java.time_with_offset_same_local",
            1,
            1,
        ),
        ("withzonesameinstant", "jvm.java.instant_with_zone", 1, 1),
        (
            "withzonesamelocal",
            "jvm.java.time_with_zone_same_local",
            1,
            1,
        ),
    ] {
        date_time.insert(name.to_string(), common_method(emit, min_args, max_args));
    }
    for type_path in [
        "time.localdatetime",
        "time.offsetdatetime",
        "time.zoneddatetime",
    ] {
        ensure_type_node(root, type_path);
        merge_type_methods(root, type_path, date_time.clone());
        merge_type_member_returns(
            root,
            type_path,
            &[
                ("plusDays", "java.time.LocalDateTime"),
                ("minusDays", "java.time.LocalDateTime"),
                ("plusWeeks", "java.time.LocalDateTime"),
                ("plusMonths", "java.time.LocalDateTime"),
                ("minusMonths", "java.time.LocalDateTime"),
                ("plusHours", "java.time.LocalDateTime"),
                ("minusHours", "java.time.LocalDateTime"),
                ("plusMinutes", "java.time.LocalDateTime"),
                ("minusMinutes", "java.time.LocalDateTime"),
                ("plusSeconds", "java.time.LocalDateTime"),
                ("minusSeconds", "java.time.LocalDateTime"),
                ("with", "java.time.LocalDateTime"),
                ("toLocalDate", "java.time.LocalDate"),
                ("toLocalTime", "java.time.LocalTime"),
                ("offset", "java.time.ZoneOffset"),
                ("zone", "java.time.ZoneId"),
                ("dayOfWeek", "java.time.DayOfWeek"),
                ("getDayOfWeek", "java.time.DayOfWeek"),
                ("toLocalDateTime", "java.time.LocalDateTime"),
                ("toInstant", "java.time.Instant"),
                ("toOffsetDateTime", "java.time.OffsetDateTime"),
                ("toZonedDateTime", "java.time.ZonedDateTime"),
            ],
        );
    }

    let mut period = Subtree::new();
    for (name, emit) in [
        ("days", "jvm.java.period_get_days"),
        ("months", "jvm.java.period_get_months"),
        ("years", "jvm.java.period_get_years"),
    ] {
        period.insert(name.to_string(), prop(emit));
    }
    for (name, emit) in [
        ("getdays", "jvm.java.period_get_days"),
        ("getmonths", "jvm.java.period_get_months"),
        ("getyears", "jvm.java.period_get_years"),
    ] {
        period.insert(name.to_string(), common_method(emit, 0, 0));
    }
    ensure_type_node(root, "time.period");
    merge_type_methods(root, "time.period", period);

    let mut year_month = Subtree::new();
    for (name, emit) in [
        ("year", "jvm.java.instant_get_year"),
        ("monthvalue", "jvm.java.instant_get_month"),
    ] {
        year_month.insert(name.to_string(), prop(emit));
    }
    year_month.insert(
        "plusmonths".to_string(),
        common_method("jvm.java.time_plus_months", 1, 1),
    );
    ensure_type_node(root, "time.yearmonth");
    merge_type_methods(root, "time.yearmonth", year_month);
    merge_type_member_returns(
        root,
        "time.yearmonth",
        &[("plusMonths", "java.time.YearMonth")],
    );

    let mut day_of_week = Subtree::new();
    day_of_week.insert("value".to_string(), prop("jvm.java.identity"));
    day_of_week.insert("name".to_string(), prop("jvm.java.day_of_week_name"));
    day_of_week.insert(
        "getvalue".to_string(),
        common_method("jvm.java.identity", 0, 0),
    );
    day_of_week.insert(
        "tostring".to_string(),
        common_method("jvm.java.day_of_week_name", 0, 0),
    );
    ensure_type_node(root, "time.dayofweek");
    merge_type_methods(root, "time.dayofweek", day_of_week);

    let mut zone_offset = Subtree::new();
    zone_offset.insert("id".to_string(), prop("jvm.java.zone_offset_id"));
    zone_offset.insert(
        "getid".to_string(),
        common_method("jvm.java.zone_offset_id", 0, 0),
    );
    ensure_type_node(root, "time.zoneoffset");
    merge_type_methods(root, "time.zoneoffset", zone_offset);

    let mut zone_id = Subtree::new();
    zone_id.insert("id".to_string(), prop("jvm.java.identity"));
    zone_id.insert(
        "getid".to_string(),
        common_method("jvm.java.identity", 0, 0),
    );
    ensure_type_node(root, "time.zoneid");
    merge_type_methods(root, "time.zoneid", zone_id);

    let mut month_day = Subtree::new();
    for (name, emit) in [
        ("monthvalue", "jvm.java.instant_get_month"),
        ("dayofmonth", "jvm.java.instant_get_day"),
    ] {
        month_day.insert(name.to_string(), prop(emit));
    }
    ensure_type_node(root, "time.monthday");
    merge_type_methods(root, "time.monthday", month_day);
}

fn insert_java_stream_statics(root: &mut Subtree) {
    for type_path in [
        "util.stream.intstream",
        "util.stream.longstream",
        "util.stream.doublestream",
        "util.stream.stream",
    ] {
        insert_common_static(root, type_path, "empty", "java.stream_empty");
        insert_common_static(root, type_path, "of", "java.stream_of");
        insert_common_static(root, type_path, "concat", "java.stream_concat");
        insert_common_static(root, type_path, "generate", "java.stream_generate");
        insert_common_static(root, type_path, "builder", "java.stream_builder");
    }
    for type_path in ["util.stream.intstream", "util.stream.longstream"] {
        insert_common_static(root, type_path, "range", "java.stream_range");
        insert_common_static(root, type_path, "rangeclosed", "java.stream_range_closed");
        insert_common_static(root, type_path, "iterate", "java.stream_iterate");
    }
    insert_common_static(
        root,
        "util.stream.doublestream",
        "iterate",
        "java.stream_iterate",
    );
    insert_common_static(
        root,
        "util.stream.stream",
        "iterate",
        "java.stream_iterate_strict",
    );

    for (member, emit) in [
        ("joining", "java.collectors_joining"),
        ("tolist", "java.collectors_to_list"),
        ("toset", "java.collectors_to_set"),
        ("tounmodifiablelist", "java.collectors_to_list"),
        ("tounmodifiableset", "java.collectors_to_set"),
        ("tocollection", "java.collectors_to_collection"),
        ("counting", "java.collectors_counting"),
        ("summingint", "java.collectors_summing_int"),
        ("summinglong", "java.collectors_summing_int"),
        ("summingdouble", "java.collectors_summing_int"),
        ("averagingint", "java.collectors_averaging_int"),
        ("averaginglong", "java.collectors_averaging_int"),
        ("averagingdouble", "java.collectors_averaging_int"),
        ("tomap", "java.collectors_to_map"),
        ("tounmodifiablemap", "java.collectors_to_map"),
        ("mapping", "java.collectors_mapping"),
        ("filtering", "java.collectors_filtering"),
        ("collectingandthen", "java.collectors_collecting_and_then"),
        ("reducing", "java.collectors_reducing"),
        ("groupingby", "java.collectors_grouping_by"),
        ("partitioningby", "java.collectors_partitioning_by"),
        ("minby", "java.collectors_min_by"),
        ("maxby", "java.collectors_max_by"),
    ] {
        insert_common_static(root, "util.stream.collectors", member, emit);
    }
}

fn insert_java_util_uuid(root: &mut Subtree) {
    let mut methods = Subtree::new();
    for (name, emit) in [
        ("version", "jvm.java.uuid_version"),
        ("variant", "jvm.java.uuid_variant"),
        ("getmostsignificantbits", "jvm.java.uuid_most_bits"),
        ("getleastsignificantbits", "jvm.java.uuid_least_bits"),
        ("hashcode", "jvm.java.uuid_hash_code"),
    ] {
        methods.insert(name.to_string(), common_emit(emit));
    }
    methods.insert(
        "compareto".to_string(),
        namespaces::overloads(vec![(1, common_emit("jvm.java.uuid_compare_to"))]),
    );
    ensure_type_node(root, "util.uuid");
    merge_type_methods(root, "util.uuid", methods);
}

fn insert_java_util_random(root: &mut Subtree) {
    let mut methods = Subtree::new();
    for (name, emit, min_args, max_args) in [
        ("setseed", "jvm.java.random_set_seed", 1, 1),
        ("nextint", "jvm.java.random_next_int", 0, 2),
        ("nextlong", "jvm.java.random_next_long", 0, 0),
        ("nextdouble", "jvm.java.random_next_double", 0, 0),
        ("nextfloat", "jvm.java.random_next_float", 0, 0),
        ("nextboolean", "jvm.java.random_next_boolean", 0, 0),
        ("nextgaussian", "jvm.java.random_next_double", 0, 0),
        ("split", "jvm.java.random_split", 0, 0),
        ("ints", "jvm.java.random_ints", 0, 3),
        ("longs", "jvm.java.random_longs", 0, 3),
        ("doubles", "jvm.java.random_doubles", 0, 3),
        ("nextbytes", "jvm.java.random_next_bytes", 1, 1),
    ] {
        methods.insert(name.to_string(), common_method(emit, min_args, max_args));
    }
    for type_path in ["util.random", "util.splittablerandom"] {
        ensure_type_node(root, type_path);
        merge_type_methods(root, type_path, methods.clone());
        merge_type_member_returns(root, type_path, &[("split", "java.util.SplittableRandom")]);
    }
}

fn insert_java_text_buffer_methods(root: &mut Subtree) {
    let mut builder = Subtree::new();
    builder.insert(
        "append".to_string(),
        common_method("jvm.java.sb_append", 1, 1),
    );
    builder.insert(
        "appendline".to_string(),
        common_method("jvm.java.sb_append_line", 0, 1),
    );
    builder.insert(
        "tostring".to_string(),
        common_method("jvm.java.sb_to_string", 0, 0),
    );
    for (name, emit, min_args, max_args) in [
        ("length", "jvm.java.sb_length", 0, 0),
        ("count", "jvm.java.sb_length", 0, 0),
        ("capacity", "jvm.java.sb_length", 0, 0),
        ("insert", "jvm.java.sb_insert", 2, 2),
        ("delete", "jvm.java.sb_delete", 2, 2),
        ("deleteat", "jvm.java.sb_delete", 1, 1),
        ("deletecharat", "jvm.java.sb_delete", 1, 1),
        ("reverse", "jvm.java.sb_reverse", 0, 0),
        ("setlength", "jvm.java.sb_set_length", 1, 1),
        ("clear", "jvm.java.sb_clear", 0, 0),
        ("setcharat", "jvm.java.sb_set_char_at", 2, 2),
        ("set", "jvm.java.sb_set_char_at", 2, 2),
        ("get", "jvm.java.sb_char_at", 1, 1),
        ("charat", "jvm.java.sb_char_at", 1, 1),
        ("appendcodepoint", "jvm.java.sb_append_code_point", 1, 1),
        ("substring", "jvm.java.sb_substring", 1, 2),
        ("subsequence", "jvm.java.sb_substring", 2, 2),
        ("replace", "jvm.java.sb_replace", 3, 3),
        ("isempty", "jvm.java.sb_is_empty", 0, 0),
        ("isnotempty", "jvm.java.sb_is_not_empty", 0, 0),
    ] {
        builder.insert(name.to_string(), common_method(emit, min_args, max_args));
    }
    ensure_type_node(root, "lang.stringbuilder");
    merge_type_methods(root, "lang.stringbuilder", builder);
    merge_type_member_returns(
        root,
        "lang.stringbuilder",
        &[
            ("append", "java.lang.StringBuilder"),
            ("appendLine", "java.lang.StringBuilder"),
            ("toString", "java.lang.String"),
        ],
    );

    let mut tokenizer = Subtree::new();
    for (name, emit) in [
        ("hasmoretokens", "jvm.java.st_has_more"),
        ("hasmoreelements", "jvm.java.st_has_more"),
        ("nexttoken", "jvm.java.st_next"),
        ("nextelement", "jvm.java.st_next"),
        ("counttokens", "jvm.java.st_count"),
    ] {
        tokenizer.insert(name.to_string(), common_method(emit, 0, 0));
    }
    ensure_type_node(root, "util.stringtokenizer");
    merge_type_methods(root, "util.stringtokenizer", tokenizer);
    merge_type_member_returns(
        root,
        "util.stringtokenizer",
        &[
            ("nextToken", "java.lang.String"),
            ("nextElement", "java.lang.String"),
        ],
    );
}

fn insert_java_net_url_uri(root: &mut Subtree) {
    let mut url_methods = Subtree::new();
    for (name, emit) in [
        ("getprotocol", "jvm.java.net.url_protocol"),
        ("gethost", "jvm.java.net.url_host"),
        ("getport", "jvm.java.net.url_port"),
        ("getdefaultport", "jvm.java.net.url_default_port"),
        ("getpath", "jvm.java.net.url_path"),
        ("getrawpath", "jvm.java.net.url_path"),
        ("getquery", "jvm.java.net.url_query"),
        ("getrawquery", "jvm.java.net.url_query"),
        ("getref", "jvm.java.net.url_ref"),
        ("getfragment", "jvm.java.net.url_ref"),
        ("getrawfragment", "jvm.java.net.url_ref"),
        ("getauthority", "jvm.java.net.url_authority"),
        ("getrawauthority", "jvm.java.net.url_authority"),
        ("getuserinfo", "jvm.java.net.url_user_info"),
        ("getrawuserinfo", "jvm.java.net.url_user_info"),
        ("tostring", "jvm.java.net.url_to_string"),
        ("toexternalform", "jvm.java.net.url_to_string"),
        ("touri", "jvm.java.net.url_to_uri"),
        ("hashcode", "jvm.java.net.url_hash"),
    ] {
        url_methods.insert(name.to_string(), common_emit(emit));
    }
    url_methods.insert(
        "equals".to_string(),
        namespaces::overloads(vec![(1, common_emit("jvm.java.net.url_equals"))]),
    );
    url_methods.insert(
        "samefile".to_string(),
        namespaces::overloads(vec![(1, common_emit("jvm.java.net.url_same_file"))]),
    );
    for (name, emit) in [
        ("protocol", "jvm.java.net.url_protocol"),
        ("host", "jvm.java.net.url_host"),
        ("port", "jvm.java.net.url_port"),
        ("defaultport", "jvm.java.net.url_default_port"),
        ("path", "jvm.java.net.url_path"),
        ("query", "jvm.java.net.url_query"),
        ("ref", "jvm.java.net.url_ref"),
        ("fragment", "jvm.java.net.url_ref"),
        ("authority", "jvm.java.net.url_authority"),
        ("file", "jvm.java.net.url_file"),
        ("userinfo", "jvm.java.net.url_user_info"),
    ] {
        url_methods.insert(name.to_string(), common_emit(emit));
    }

    let mut url_returns = BTreeMap::new();
    url_returns.insert("touri".to_string(), "java.net.URI".to_string());

    insert_path(
        root,
        "net.url",
        NamespaceNode::Type {
            ctor: Some(namespaces::CtorSpec {
                params: Vec::new(),
                fields: Vec::new(),
                ancestry: ["URL", "Serializable", "Object"]
                    .iter()
                    .map(|s| (*s).to_string())
                    .collect(),
                control_fn: None,
                nest_coerce: None,
                field_gui: Vec::new(),
                value_equality: false,
            }),
            ctor_call: Some(Box::new(common_emit("jvm.java.net.url_new"))),
            statics: Subtree::new(),
            methods: url_methods,
            member_returns: url_returns,
        },
    );

    let mut encoder_statics = Subtree::new();
    encoder_statics.insert("encode".to_string(), common_emit("jvm.java.net.url_encode"));
    insert_path(
        root,
        "net.urlencoder",
        NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics: encoder_statics,
            methods: Subtree::new(),
            member_returns: Default::default(),
        },
    );

    let mut decoder_statics = Subtree::new();
    decoder_statics.insert("decode".to_string(), common_emit("jvm.java.net.url_decode"));
    insert_path(
        root,
        "net.urldecoder",
        NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics: decoder_statics,
            methods: Subtree::new(),
            member_returns: Default::default(),
        },
    );

    let mut uri_statics = Subtree::new();
    uri_statics.insert("create".to_string(), common_emit("jvm.java.net.uri_new"));
    let mut uri_methods = Subtree::new();
    for (name, emit) in [
        ("getscheme", "jvm.java.net.url_protocol"),
        ("gethost", "jvm.java.net.url_host"),
        ("getport", "jvm.java.net.url_port"),
        ("getpath", "jvm.java.net.url_path"),
        ("getrawpath", "jvm.java.net.url_path"),
        ("getquery", "jvm.java.net.url_query"),
        ("getrawquery", "jvm.java.net.url_query"),
        ("getfragment", "jvm.java.net.url_ref"),
        ("getrawfragment", "jvm.java.net.url_ref"),
        ("getauthority", "jvm.java.net.url_authority"),
        ("getrawauthority", "jvm.java.net.url_authority"),
        ("getuserinfo", "jvm.java.net.url_user_info"),
        ("getrawuserinfo", "jvm.java.net.url_user_info"),
        ("getschemespecificpart", "jvm.java.net.uri_ssp"),
        ("getrawschemespecificpart", "jvm.java.net.uri_ssp"),
        ("isabsolute", "jvm.java.net.uri_is_absolute"),
        ("isopaque", "jvm.java.net.uri_is_opaque"),
        ("normalize", "jvm.java.net.uri_normalize"),
        ("tostring", "jvm.java.net.url_to_string"),
        ("toasciistring", "jvm.java.net.url_to_string"),
        ("tourl", "jvm.java.net.uri_to_url"),
        ("hashcode", "jvm.java.net.url_hash"),
    ] {
        uri_methods.insert(name.to_string(), common_emit(emit));
    }
    for (name, emit) in [
        ("resolve", "jvm.java.net.uri_resolve"),
        ("relativize", "jvm.java.net.uri_relativize"),
        ("compareto", "jvm.java.net.uri_compare_to"),
        ("equals", "jvm.java.net.url_equals"),
    ] {
        uri_methods.insert(
            name.to_string(),
            namespaces::overloads(vec![(1, common_emit(emit))]),
        );
    }
    for (name, emit) in [
        ("scheme", "jvm.java.net.url_protocol"),
        ("host", "jvm.java.net.url_host"),
        ("port", "jvm.java.net.url_port"),
        ("path", "jvm.java.net.url_path"),
        ("query", "jvm.java.net.url_query"),
        ("fragment", "jvm.java.net.url_ref"),
        ("authority", "jvm.java.net.url_authority"),
        ("userinfo", "jvm.java.net.url_user_info"),
        ("schemespecificpart", "jvm.java.net.uri_ssp"),
        ("isabsolute", "jvm.java.net.uri_is_absolute"),
        ("isopaque", "jvm.java.net.uri_is_opaque"),
    ] {
        uri_methods.insert(name.to_string(), common_emit(emit));
    }

    let mut uri_returns = BTreeMap::new();
    for name in ["normalize", "resolve", "relativize"] {
        uri_returns.insert(name.to_string(), "java.net.URI".to_string());
    }
    uri_returns.insert("tourl".to_string(), "java.net.URL".to_string());

    insert_path(
        root,
        "net.uri",
        NamespaceNode::Type {
            ctor: Some(namespaces::CtorSpec {
                params: Vec::new(),
                fields: Vec::new(),
                ancestry: ["URI", "Comparable", "Serializable", "Object"]
                    .iter()
                    .map(|s| (*s).to_string())
                    .collect(),
                control_fn: None,
                nest_coerce: None,
                field_gui: Vec::new(),
                value_equality: false,
            }),
            ctor_call: Some(Box::new(common_emit("jvm.java.net.uri_new"))),
            statics: uri_statics,
            methods: uri_methods,
            member_returns: uri_returns,
        },
    );
}

/// One Java stdlib type, as DATA: what it is called, what it IS (its
/// `is`/`isInstance`/`instanceof` ancestry — self first, then supertypes and
/// interfaces), and — when the type has no object identity to carry a stamp —
/// which runtime intrinsic backs it.
pub struct JavaType {
    pub name: &'static str,
    /// The package the type lives in. Java packages ARE its namespaces, so
    /// this is where the type registers in the tree — `java.util.arraylist`.
    /// One registration then answers both `ArrayList` (matched on the leaf)
    /// and `java.util.ArrayList` (matched on the full path).
    pub package: &'static str,
    pub ancestry: &'static [&'static str],
    /// The SHARED canonical type name that `ExprKind::IsType` already tests
    /// for (`"string"`, `"number"`, `"boolean"`, `"list"`, `"object"`).
    ///
    /// `Some` means the type's runtime representation is an intrinsic, so it
    /// cannot carry a `__type`/`__types` stamp and identity has to be answered
    /// by testing the value's runtime kind. `None` means the value is an
    /// object and answers from the stamped ancestry chain like any user class.
    ///
    /// This is the ONLY Java-specific part of the identity story, and it is a
    /// name, not an emit: the test itself is the shared one.
    pub intrinsic: Option<&'static str>,
}

const fn t(
    name: &'static str,
    package: &'static str,
    ancestry: &'static [&'static str],
    intrinsic: Option<&'static str>,
) -> JavaType {
    JavaType {
        name,
        package,
        ancestry,
        intrinsic,
    }
}

/// The Java types whose IDENTITY the runtime must answer for.
///
/// These are DATA, and they are Java knowledge, so they live in the Java crate.
/// What they are NOT is a second resolution path: registering them as `Type`
/// nodes puts them on the same common-resolver `Ctor` + `ancestry` machinery
/// that dotnet, flutter and plib use, and the `intrinsic` column feeds the
/// walker's normalisation of `X.class.isInstance(v)` / `v instanceof X` into
/// the shared `ExprKind::IsType` — so no Java-specific identity check emits.
///
/// Before this they were registered as bare `CommonEmit` leaves — callable, but
/// with no identity — so a type used as a VALUE (`X.class`) resolved to nothing
/// and the call trapped with "undefined is not callable". The same shape as a
/// platform class that was never materialised.
pub const JAVA_TYPES: &[JavaType] = &[
    // ── Intrinsic-backed: no object to stamp, so identity is a kind test ──
    t(
        "String",
        "lang",
        &[
            "String",
            "CharSequence",
            "Comparable",
            "Serializable",
            "Object",
        ],
        Some("string"),
    ),
    t(
        "Character",
        "lang",
        &["Character", "Comparable", "Serializable", "Object"],
        Some("string"),
    ),
    t(
        "Class",
        "lang",
        &["Class", "Serializable", "Object"],
        Some("string"),
    ),
    t(
        "Integer",
        "lang",
        &["Integer", "Number", "Comparable", "Serializable", "Object"],
        Some("number"),
    ),
    t(
        "Long",
        "lang",
        &["Long", "Number", "Comparable", "Serializable", "Object"],
        Some("number"),
    ),
    t(
        "Short",
        "lang",
        &["Short", "Number", "Comparable", "Serializable", "Object"],
        Some("number"),
    ),
    t(
        "Byte",
        "lang",
        &["Byte", "Number", "Comparable", "Serializable", "Object"],
        Some("number"),
    ),
    t(
        "Double",
        "lang",
        &["Double", "Number", "Comparable", "Serializable", "Object"],
        Some("number"),
    ),
    t(
        "Float",
        "lang",
        &["Float", "Number", "Comparable", "Serializable", "Object"],
        Some("number"),
    ),
    t(
        "Boolean",
        "lang",
        &["Boolean", "Comparable", "Serializable", "Object"],
        Some("boolean"),
    ),
    // The root: its own ancestry is just itself, so it contributes `"object"`
    // to `Object.class.isInstance(…)` without leaking that kind into every
    // interface query. Java's `new Object()` is a bare map at runtime.
    t("Object", "lang", &["Object"], None),
    // ── Object-backed: identity comes from the stamped ancestry chain ──
    t(
        "ArrayList",
        "util",
        &[
            "ArrayList",
            "List",
            "Collection",
            "Iterable",
            "Cloneable",
            "Object",
        ],
        None,
    ),
    t(
        "LinkedList",
        "util",
        &[
            "LinkedList",
            "Deque",
            "Queue",
            "List",
            "Collection",
            "Iterable",
            "Cloneable",
            "Object",
        ],
        None,
    ),
    t(
        "ArrayDeque",
        "util",
        &[
            "ArrayDeque",
            "Deque",
            "Queue",
            "Collection",
            "Iterable",
            "Cloneable",
            "Object",
        ],
        None,
    ),
    t(
        "Vector",
        "util",
        &[
            "Vector",
            "List",
            "Collection",
            "Iterable",
            "Cloneable",
            "Object",
        ],
        None,
    ),
    t(
        "Stack",
        "util",
        &[
            "Stack",
            "Vector",
            "List",
            "Collection",
            "Iterable",
            "Cloneable",
            "Object",
        ],
        None,
    ),
    t(
        "HashSet",
        "util",
        &[
            "HashSet",
            "Set",
            "Collection",
            "Iterable",
            "Cloneable",
            "Object",
        ],
        None,
    ),
    t(
        "LinkedHashSet",
        "util",
        &[
            "LinkedHashSet",
            "HashSet",
            "Set",
            "Collection",
            "Iterable",
            "Cloneable",
            "Object",
        ],
        None,
    ),
    t(
        "TreeSet",
        "util",
        &[
            "TreeSet",
            "NavigableSet",
            "SortedSet",
            "Set",
            "Collection",
            "Iterable",
            "Object",
        ],
        None,
    ),
    t(
        "PriorityQueue",
        "util",
        &["PriorityQueue", "Queue", "Collection", "Iterable", "Object"],
        None,
    ),
    t("Comparator", "util", &["Comparator", "Object"], None),
    t("Iterator", "util", &["Iterator", "Object"], None),
    t(
        "ListIterator",
        "util",
        &["ListIterator", "Iterator", "Object"],
        None,
    ),
    t(
        "HashMap",
        "util",
        &["HashMap", "Map", "Cloneable", "Object"],
        None,
    ),
    t(
        "LinkedHashMap",
        "util",
        &["LinkedHashMap", "HashMap", "Map", "Cloneable", "Object"],
        None,
    ),
    t(
        "TreeMap",
        "util",
        &["TreeMap", "NavigableMap", "SortedMap", "Map", "Object"],
        None,
    ),
    t(
        "StringBuilder",
        "lang",
        &["StringBuilder", "CharSequence", "Appendable", "Object"],
        None,
    ),
    t(
        "StringBuffer",
        "lang",
        &["StringBuffer", "CharSequence", "Appendable", "Object"],
        None,
    ),
    t(
        "StringTokenizer",
        "util",
        &["StringTokenizer", "Enumeration", "Object"],
        None,
    ),
    // Types whose `[known_types]` entry declares a constructor but which had no
    // JAVA_TYPES row, so they never registered as tree `Type` nodes. While the
    // declarations lived in the Java profile that was invisible —
    // `lookup_known_type` answered from the language's own table. Once the
    // qualified names moved here, the TREE is the only thing that can answer,
    // so the row is what makes `new java.math.BigInteger(...)` resolve at all.
    t(
        "BigInteger",
        "math",
        &[
            "BigInteger",
            "Number",
            "Comparable",
            "Serializable",
            "Object",
        ],
        None,
    ),
    t(
        "BitSet",
        "util",
        &["BitSet", "Cloneable", "Serializable", "Object"],
        None,
    ),
    t(
        "UUID",
        "util",
        &["UUID", "Comparable", "Serializable", "Object"],
        None,
    ),
    t(
        "Random",
        "util",
        &["Random", "Serializable", "Object"],
        None,
    ),
    t(
        "SplittableRandom",
        "util",
        &["SplittableRandom", "Object"],
        None,
    ),
    t(
        "Pattern",
        "util.regex",
        &["Pattern", "Object"],
        None,
    ),
    t(
        "Matcher",
        "util.regex",
        &["Matcher", "Object"],
        None,
    ),
    t(
        "ByteArrayOutputStream",
        "io",
        &[
            "ByteArrayOutputStream",
            "OutputStream",
            "Closeable",
            "AutoCloseable",
            "Object",
        ],
        None,
    ),
    t(
        "ByteArrayInputStream",
        "io",
        &[
            "ByteArrayInputStream",
            "InputStream",
            "Closeable",
            "AutoCloseable",
            "Object",
        ],
        None,
    ),
    t(
        "File",
        "io",
        &["File", "Serializable", "Comparable", "Object"],
        None,
    ),
    t(
        "PrintWriter",
        "io",
        &[
            "PrintWriter",
            "Writer",
            "Closeable",
            "Flushable",
            "AutoCloseable",
            "Object",
        ],
        None,
    ),
    t(
        "PrintStream",
        "io",
        &[
            "PrintStream",
            "OutputStream",
            "Appendable",
            "Closeable",
            "Flushable",
            "AutoCloseable",
            "Object",
        ],
        None,
    ),
    t(
        "OutputStreamWriter",
        "io",
        &[
            "OutputStreamWriter",
            "Writer",
            "Closeable",
            "Flushable",
            "AutoCloseable",
            "Object",
        ],
        None,
    ),
    t(
        "InputStreamReader",
        "io",
        &[
            "InputStreamReader",
            "Reader",
            "Closeable",
            "AutoCloseable",
            "Object",
        ],
        None,
    ),
    t(
        "BufferedReader",
        "io",
        &[
            "BufferedReader",
            "Reader",
            "Closeable",
            "AutoCloseable",
            "Object",
        ],
        None,
    ),
    t(
        "BufferedWriter",
        "io",
        &[
            "BufferedWriter",
            "Writer",
            "Closeable",
            "Flushable",
            "AutoCloseable",
            "Object",
        ],
        None,
    ),
    t(
        "StringReader",
        "io",
        &[
            "StringReader",
            "Reader",
            "Closeable",
            "AutoCloseable",
            "Object",
        ],
        None,
    ),
    t(
        "StringWriter",
        "io",
        &[
            "StringWriter",
            "Writer",
            "Appendable",
            "Closeable",
            "Flushable",
            "AutoCloseable",
            "Object",
        ],
        None,
    ),
    t(
        "CharArrayReader",
        "io",
        &[
            "CharArrayReader",
            "Reader",
            "Closeable",
            "AutoCloseable",
            "Object",
        ],
        None,
    ),
    t(
        "CharArrayWriter",
        "io",
        &[
            "CharArrayWriter",
            "Writer",
            "Appendable",
            "Closeable",
            "Flushable",
            "AutoCloseable",
            "Object",
        ],
        None,
    ),
    t(
        "PushbackInputStream",
        "io",
        &[
            "PushbackInputStream",
            "FilterInputStream",
            "InputStream",
            "Closeable",
            "AutoCloseable",
            "Object",
        ],
        None,
    ),
    t(
        "BufferedInputStream",
        "io",
        &[
            "BufferedInputStream",
            "FilterInputStream",
            "InputStream",
            "Closeable",
            "AutoCloseable",
            "Object",
        ],
        None,
    ),
    t(
        "FilterInputStream",
        "io",
        &[
            "FilterInputStream",
            "InputStream",
            "Closeable",
            "AutoCloseable",
            "Object",
        ],
        None,
    ),
    t(
        "FilterWriter",
        "io",
        &[
            "FilterWriter",
            "Writer",
            "Closeable",
            "Flushable",
            "AutoCloseable",
            "Object",
        ],
        None,
    ),
    t(
        "LineNumberReader",
        "io",
        &[
            "LineNumberReader",
            "BufferedReader",
            "Reader",
            "Closeable",
            "AutoCloseable",
            "Object",
        ],
        None,
    ),
    t(
        "SequenceInputStream",
        "io",
        &[
            "SequenceInputStream",
            "InputStream",
            "Closeable",
            "AutoCloseable",
            "Object",
        ],
        None,
    ),
    t(
        "DataOutputStream",
        "io",
        &[
            "DataOutputStream",
            "FilterOutputStream",
            "OutputStream",
            "Closeable",
            "Flushable",
            "AutoCloseable",
            "Object",
        ],
        None,
    ),
    t(
        "DataInputStream",
        "io",
        &[
            "DataInputStream",
            "FilterInputStream",
            "InputStream",
            "Closeable",
            "AutoCloseable",
            "Object",
        ],
        None,
    ),
    t(
        "Hashtable",
        "util",
        &["Hashtable", "Map", "Cloneable", "Serializable", "Object"],
        None,
    ),
    t(
        "IdentityHashMap",
        "util",
        &[
            "IdentityHashMap",
            "Map",
            "Cloneable",
            "Serializable",
            "Object",
        ],
        None,
    ),
    t(
        "Properties",
        "util",
        &["Properties", "Hashtable", "Map", "Cloneable", "Object"],
        None,
    ),
    t(
        "WeakHashMap",
        "util",
        &["WeakHashMap", "Map", "Object"],
        None,
    ),
    t(
        "CopyOnWriteArrayList",
        "util.concurrent",
        &[
            "CopyOnWriteArrayList",
            "List",
            "Collection",
            "Iterable",
            "Cloneable",
            "Object",
        ],
        None,
    ),
    t(
        "LinkedBlockingQueue",
        "util.concurrent",
        &[
            "LinkedBlockingQueue",
            "BlockingQueue",
            "Queue",
            "Collection",
            "Iterable",
            "Object",
        ],
        None,
    ),
    t(
        "Semaphore",
        "util.concurrent",
        &["Semaphore", "Serializable", "Object"],
        None,
    ),
    // The one Throwable the profile constructs as a plain map rather than
    // through `ecma:error`. A user class that EXTENDS a built-in exception
    // already gets its chain stamped by the walker; this is the direct
    // `new OutOfMemoryError()` case, which had no identity at all.
    t(
        "OutOfMemoryError",
        "lang",
        &[
            "OutOfMemoryError",
            "VirtualMachineError",
            "Error",
            "Throwable",
            "Object",
        ],
        None,
    ),
];

/// A Java array type and the runtime kind of its ELEMENTS.
///
/// Every Java array is a JS array at runtime and the element type is not
/// represented anywhere on the value, so `String[]` and `int[]` are the same
/// object. The only thing that distinguishes them is what is IN them — hence
/// the element kind, which the caller probes on the first element.
///
/// An array type not listed here (`Object[]`, a generic `T[]`) answers on
/// array-ness alone, which is the most that can be said about it.
pub fn array_element_intrinsic(type_name: &str) -> Option<&'static str> {
    let element = type_name.strip_suffix("[]")?.trim();
    Some(match element {
        "int" | "long" | "short" | "byte" | "double" | "float" => "number",
        "char" | "String" => "string",
        "boolean" => "boolean",
        _ => return None,
    })
}

/// Every distinct runtime intrinsic that could answer `isInstance` for
/// `queried` — the intrinsics of all registered types that HAVE `queried` in
/// their ancestry. `Comparable` yields `["string", "number", "boolean"]`;
/// `String` yields `["string"]`; a user class yields nothing.
///
/// Empty means "no intrinsic can answer this" — the caller falls back to the
/// stamped-ancestry path, which is also what must be OR-ed in alongside a
/// non-empty result so object-backed instances still answer.
pub fn intrinsics_answering(queried: &str) -> Vec<&'static str> {
    let mut out: Vec<&'static str> = Vec::new();
    for ty in JAVA_TYPES {
        let Some(intrinsic) = ty.intrinsic else {
            continue;
        };
        if ty.ancestry.iter().any(|a| *a == queried) && !out.contains(&intrinsic) {
            out.push(intrinsic);
        }
    }
    out
}

/// Register the JVM stdlib surface under platform-owned roots. Idempotent;
/// first call wins.
pub fn register_namespace_tree() {
    static ONCE: Once = Once::new();
    ONCE.call_once(|| {
        let mut root = Subtree::new();

        // TYPES first, at their package path (`java.util.arraylist`), so one
        // registration answers both the bare name and the fully-qualified
        // chain. Constructor targets are platform tree data, not language
        // profile rows; the `Type` node wraps them so construction also STAMPS
        // the declared ancestry and `isInstance` can answer from `__types`.
        //
        // Only object-backed types register. An intrinsic-backed type
        // (`String`, `Integer`) has no object to stamp, so a `Type` node would
        // promise an identity it cannot carry; those answer through the
        // `intrinsic` column instead.
        for ty in JAVA_TYPES.iter().filter(|ty| ty.intrinsic.is_none()) {
            // The QUALIFIED key: this fragment owns `java.util.ArrayList`, not
            // the bare `ArrayList` spelling — that is Java's own idiom and
            // stays in the Java profile. A platform declares package paths.
            let qualified = format!("java.{}.{}", ty.package, ty.name);
            let Some(ctor_call) = java_type_ctor_target(&qualified) else {
                continue;
            };
            let path = format!("{}.{}", ty.package, ty.name.to_lowercase());
            insert_path(
                &mut root,
                &path,
                NamespaceNode::Type {
                    ctor: Some(namespaces::CtorSpec {
                        params: Vec::new(),
                        fields: Vec::new(),
                        ancestry: ty.ancestry.iter().map(|s| (*s).to_string()).collect(),
                        control_fn: None,
                        nest_coerce: None,
                        field_gui: Vec::new(),
                        value_equality: false,
                    }),
                    ctor_call: Some(Box::new(ctor_call)),
                    statics: Subtree::new(),
                    methods: Subtree::new(),
                    member_returns: Default::default(),
                },
            );
        }

        insert_java_lang_system(&mut root);
        insert_java_util_core_statics(&mut root);
        insert_java_lang_core_statics(&mut root);
        insert_java_time_statics(&mut root);
        insert_java_time_instance_members(&mut root);
        insert_java_stream_statics(&mut root);
        insert_java_net_url_uri(&mut root);
        insert_java_util_collection_methods(&mut root);
        insert_java_util_regex(&mut root);
        insert_java_math_biginteger_methods(&mut root);
        insert_java_util_collection_statics(&mut root);
        insert_java_util_enum_set(&mut root);
        insert_java_io_methods(&mut root);
        insert_java_util_uuid(&mut root);
        insert_java_util_random(&mut root);
        insert_java_text_buffer_methods(&mut root);
        insert_java_namespace_constants(&mut root);
        let mut jvm_root = Subtree::new();
        jvm_root.insert("java".to_string(), NamespaceNode::Namespace(root));
        namespaces::register_namespace_tree("jvm", NamespaceNode::Namespace(jvm_root));
    });
}

#[cfg(test)]
mod tests {
    #[test]
    fn profile_fragment_parses_and_declares_the_surface() {
        let profile = vybe_runtime::profile::parse_profile(crate::profile_source())
            .expect("java.* profile fragment must parse");
        assert!(
            profile.known_types.is_empty(),
            "platform constructors are tree leaves, not profile known_types"
        );
        assert!(
            profile.namespace_constants.is_empty(),
            "platform constants are tree leaves, not profile namespace_constants"
        );
        assert!(
            profile.value_methods.is_empty(),
            "platform instance methods are tree leaves, not profile value_methods"
        );
        assert!(
            profile.builtins.is_empty(),
            "platform builtins are tree leaves, not profile rows"
        );
    }

    #[test]
    fn java_constructors_are_tree_registered_not_profile_known_types() {
        let profile = vybe_runtime::profile::parse_profile(crate::profile_source())
            .expect("java.* profile fragment must parse");
        assert!(
            profile.lookup_known_type("java.util.ArrayList").is_none(),
            "platform constructors should be registered by tree_register.rs, not profile known_types"
        );

        super::register_namespace_tree();
        let scopes = vec!["jvm".to_string()];
        assert!(
            vybe_runtime::namespaces::lookup_type_ctor_target(&scopes, "java.util.ArrayList")
                .is_some()
        );
        assert!(
            vybe_runtime::namespaces::lookup_type_ctor_target(&scopes, "java.util.HashMap")
                .is_some()
        );
        assert!(
            vybe_runtime::namespaces::lookup_type_ctor_target(&scopes, "java.util.UUID").is_some()
        );
        assert!(
            vybe_runtime::namespaces::lookup_type_ctor_target(&scopes, "java.lang.Object")
                .is_some()
        );
        assert!(
            vybe_runtime::namespaces::lookup_type_ctor_target(&scopes, "java.lang.StringBuffer")
                .is_some()
        );
        assert!(
            vybe_runtime::namespaces::lookup_type_ctor_target(&scopes, "java.util.Random")
                .is_some()
        );
        assert!(
            vybe_runtime::namespaces::lookup_type_ctor_target(
                &scopes,
                "java.util.SplittableRandom"
            )
            .is_some()
        );
    }

    #[test]
    fn java_net_tree_declares_method_return_types() {
        super::register_namespace_tree();
        let scopes = vec!["jvm".to_string()];
        assert!(
            vybe_runtime::namespaces::lookup_type_static_member(
                &scopes,
                "java.lang.System",
                "getProperty"
            )
            .is_some()
        );
        assert_eq!(
            vybe_runtime::namespaces::lookup_type_member_return(&scopes, "URI", "resolve"),
            Some("java.net.URI".to_string())
        );
        assert_eq!(
            vybe_runtime::namespaces::lookup_type_member_return(&scopes, "java.net.URI", "toURL"),
            Some("java.net.URL".to_string())
        );
        assert_eq!(
            vybe_runtime::namespaces::lookup_type_member_return(&scopes, "URL", "toURI"),
            Some("java.net.URI".to_string())
        );
    }

    #[test]
    fn java_lang_stringbuilder_tree_is_platform_owned() {
        super::register_namespace_tree();
        let scopes = vec!["jvm".to_string()];
        assert!(
            vybe_runtime::namespaces::lookup_type_ctor_target(&scopes, "java.lang.StringBuilder")
                .is_some()
        );
        assert!(
            vybe_runtime::namespaces::lookup_type_instance_target(
                &scopes,
                "java.lang.StringBuilder",
                "append",
                1,
            )
            .is_some()
        );
        assert_eq!(
            vybe_runtime::namespaces::lookup_type_member_return(
                &scopes,
                "java.lang.StringBuilder",
                "append",
            ),
            Some("java.lang.StringBuilder".to_string())
        );
    }

    #[test]
    fn java_util_stringtokenizer_tree_is_platform_owned() {
        super::register_namespace_tree();
        let scopes = vec!["jvm".to_string()];
        assert!(
            vybe_runtime::namespaces::lookup_type_ctor_target(&scopes, "java.util.StringTokenizer")
                .is_some()
        );
        assert!(
            vybe_runtime::namespaces::lookup_type_instance_target(
                &scopes,
                "java.util.StringTokenizer",
                "hasMoreTokens",
                0,
            )
            .is_some()
        );
        assert_eq!(
            vybe_runtime::namespaces::lookup_type_member_return(
                &scopes,
                "java.util.StringTokenizer",
                "nextToken",
            ),
            Some("java.lang.String".to_string())
        );
    }

    #[test]
    fn java_util_uuid_tree_is_platform_owned() {
        super::register_namespace_tree();
        let scopes = vec!["jvm".to_string()];
        assert!(matches!(
            vybe_runtime::namespaces::lookup_type_ctor_target(&scopes, "java.util.UUID"),
            Some(vybe_runtime::component_model::ConstructorTarget::Common(op))
                if op == "jvm.java.uuid_new"
        ));
        assert!(matches!(
            vybe_runtime::namespaces::lookup_type_static_member(
                &scopes,
                "java.util.UUID",
                "fromString"
            ),
            Some(vybe_runtime::namespaces::NamespaceNode::CommonEmit(op))
                if op == "jvm.java.uuid_from_string"
        ));
        assert!(matches!(
            vybe_runtime::namespaces::lookup_type_instance_member(
                &scopes,
                "java.util.UUID",
                "version"
            ),
            Some(vybe_runtime::namespaces::NamespaceNode::CommonEmit(op))
                if op == "jvm.java.uuid_version"
        ));
    }

    #[test]
    fn java_util_objects_tree_is_platform_owned() {
        super::register_namespace_tree();
        let scopes = vec!["jvm".to_string()];
        assert!(matches!(
            vybe_runtime::namespaces::lookup_type_static_member(
                &scopes,
                "java.util.Objects",
                "equals"
            ),
            Some(vybe_runtime::namespaces::NamespaceNode::CommonEmit(op))
                if op == "object.equals"
        ));
        assert!(matches!(
            vybe_runtime::namespaces::lookup_type_static_member(
                &scopes,
                "java.util.Objects",
                "toString"
            ),
            Some(vybe_runtime::namespaces::NamespaceNode::CommonEmit(op))
                if op == "object.to_string_or"
        ));
    }

    #[test]
    fn java_lang_character_tree_is_platform_owned() {
        super::register_namespace_tree();
        let scopes = vec!["jvm".to_string()];
        assert!(matches!(
            vybe_runtime::namespaces::lookup_type_static_member(
                &scopes,
                "java.lang.Character",
                "isDigit"
            ),
            Some(vybe_runtime::namespaces::NamespaceNode::CommonEmit(op))
                if op == "jvm.java.char_is_digit"
        ));
        assert!(matches!(
            vybe_runtime::namespaces::lookup_type_static_member(
                &scopes,
                "java.lang.Character",
                "toUpperCase"
            ),
            Some(vybe_runtime::namespaces::NamespaceNode::CommonEmit(op))
                if op == "jvm.java.char_to_upper"
        ));
    }

    #[test]
    fn java_lang_integer_tree_is_platform_owned() {
        super::register_namespace_tree();
        let scopes = vec!["jvm".to_string()];
        assert!(matches!(
            vybe_runtime::namespaces::lookup_type_static_member(
                &scopes,
                "java.lang.Integer",
                "parseInt"
            ),
            Some(vybe_runtime::namespaces::NamespaceNode::CommonEmit(op))
                if op == "jvm.java.parse_int"
        ));
        assert!(matches!(
            vybe_runtime::namespaces::lookup_type_static_member(
                &scopes,
                "java.lang.Integer",
                "bitCount"
            ),
            Some(vybe_runtime::namespaces::NamespaceNode::CommonEmit(op))
                if op == "jvm.java.int_bit_count"
        ));
    }

    #[test]
    fn java_util_arrays_tree_is_platform_owned() {
        super::register_namespace_tree();
        let scopes = vec!["jvm".to_string()];
        assert!(matches!(
            vybe_runtime::namespaces::lookup_type_static_member(
                &scopes,
                "java.util.Arrays",
                "sort"
            ),
            Some(vybe_runtime::namespaces::NamespaceNode::CommonEmit(op))
                if op == "jvm.java.arrays_sort"
        ));
        assert!(matches!(
            vybe_runtime::namespaces::lookup_type_static_member(
                &scopes,
                "java.util.Arrays",
                "asList"
            ),
            Some(vybe_runtime::namespaces::NamespaceNode::CommonEmit(op))
                if op == "jvm.java.arrays_as_list"
        ));
    }

    #[test]
    fn java_util_bitset_tree_is_platform_owned() {
        super::register_namespace_tree();
        let scopes = vec!["jvm".to_string()];
        assert!(matches!(
            vybe_runtime::namespaces::lookup_type_ctor_target(&scopes, "java.util.BitSet"),
            Some(vybe_runtime::component_model::ConstructorTarget::Common(op))
                if op == "jvm.java.bitset_new"
        ));
        assert!(matches!(
            vybe_runtime::namespaces::lookup_type_static_member(
                &scopes,
                "java.util.BitSet",
                "valueOf"
            ),
            Some(vybe_runtime::namespaces::NamespaceNode::CommonEmit(op))
                if op == "jvm.java.bitset_value_of"
        ));
    }

    /// `java.util.EnumSet` is reachable from an arbitrary scope by PATH — no
    /// `common:java.enum_set_*` profile row, no walker rewrite.
    ///
    /// The statics, the instance methods and the declared return type are
    /// asserted separately on purpose: registering the leaves without
    /// promoting the node to a `Type` left the statics resolving while
    /// `find_type_node` — which matches `Type` nodes only — silently dropped
    /// every method and every `member_returns` entry.
    #[test]
    fn java_util_enum_set_tree_is_platform_owned() {
        super::register_namespace_tree();
        let scopes = vec!["jvm".to_string()];
        assert!(matches!(
            vybe_runtime::namespaces::lookup_type_static_member(
                &scopes,
                "java.util.EnumSet",
                "allOf"
            ),
            Some(node) if matches!(
                vybe_runtime::namespaces::select_overload(&node, 1),
                Some(vybe_runtime::namespaces::NamespaceNode::CommonEmit(op))
                    if op == "jvm.java.enum_set_all_of"
            )
        ));
        assert!(matches!(
            vybe_runtime::namespaces::lookup_type_instance_target(
                &scopes,
                "java.util.EnumSet",
                "contains",
                1,
            ),
            Some(vybe_runtime::component_model::InstanceMethodTarget::Common { emit, .. })
                if emit == "jvm.java.enum_set_contains"
        ));
        assert_eq!(
            vybe_runtime::namespaces::lookup_type_member_return(&scopes, "EnumSet", "of")
                .as_deref(),
            Some("java.util.EnumSet"),
        );
        // `java.lang.Enum`'s metadata hook, which is what makes a leaf handed
        // only `X.class` — a NAME — able to reach the constants.
        assert!(
            vybe_runtime::namespaces::lookup_type_static_member(
                &scopes,
                "java.lang.Enum",
                "__vybe_declare"
            )
            .is_some()
        );
    }
}
