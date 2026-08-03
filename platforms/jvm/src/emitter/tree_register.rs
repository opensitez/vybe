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
            _ => return };
    }
    let Some(existing) = cursor.get_mut(leaf) else {
        return;
    };
    if let NamespaceNode::Namespace(children) = existing {
        let statics = std::mem::take(children);
        *existing = NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics,
            methods: Subtree::new(),
            member_returns: Default::default() };
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
            _ => return };
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
            _ => return };
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
    Str(&'static str) }

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
            JavaConst::Str(v) => NamespaceNode::Const(Value::String(v.into())) };
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
        "java.util.BitSet" => "jvm.java.bitset_new",
        "java.util.UUID" => "jvm.java.uuid_new",
        "java.util.Random" | "java.util.SplittableRandom" => "jvm.java.random_new",
        "java.lang.StringBuilder" | "java.lang.StringBuffer" => "jvm.java.stringbuilder_new",
        "java.util.StringTokenizer" => "jvm.java.stringtokenizer_new",
        "java.lang.Object" => "jvm.java.hash_map_new",
        "java.io.ByteArrayOutputStream" => "jvm.java.io_byte_array_output_stream_new",
        "java.io.ByteArrayInputStream" => "jvm.java.io_byte_array_input_stream_new",
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
        _ => return None };
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
    insert_path(
        root,
        "lang.system",
        NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics,
            methods: Subtree::new(),
            member_returns: Default::default() },
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

fn insert_java_util_collection_methods(root: &mut Subtree) {
    const SPECS: &[(&str, &str, &str, u8, u8)] = &[
        ("util.arraylist", "add", "jvm.java.add", 1, 2),
        ("util.arraylist", "get", "jvm.java.get", 1, 1),
        ("util.arraylist", "set", "jvm.java.list_set", 1, 2),
        ("util.arraylist", "size", "jvm.java.size", 0, 0),
        ("util.arraylist", "clear", "jvm.java.list_clear", 0, 0),
        ("util.arraylist", "sort", "jvm.java.collections_sort", 0, 1),
        ("util.linkedlist", "add", "jvm.java.add", 1, 2),
        ("util.linkedlist", "offer", "collections.push", 1, 1),
        ("util.linkedlist", "offerlast", "collections.push", 1, 1),
        ("util.linkedlist", "addlast", "collections.push", 1, 1),
        ("util.linkedlist", "addfirst", "jvm.java.add_first", 1, 1),
        (
            "util.linkedlist",
            "removefirst",
            "jvm.java.remove_first",
            0,
            0,
        ),
        ("util.linkedlist", "removelast", "collections.pop", 0, 0),
        ("util.linkedlist", "poll", "jvm.java.queue_poll", 0, 0),
        ("util.linkedlist", "peek", "jvm.java.peek_first", 0, 0),
        ("util.linkedlist", "get", "jvm.java.get", 1, 1),
        ("util.linkedlist", "size", "jvm.java.size", 0, 0),
        ("util.vector", "add", "jvm.java.add", 1, 2),
        ("util.vector", "addelement", "jvm.java.add", 1, 1),
        ("util.vector", "get", "jvm.java.get", 1, 1),
        ("util.vector", "elementat", "jvm.java.get", 1, 1),
        ("util.vector", "iterator", "jvm.java.list_iterator", 0, 0),
        (
            "util.vector",
            "listiterator",
            "jvm.java.list_iterator",
            0,
            1,
        ),
        ("util.vector", "set", "jvm.java.list_set", 1, 2),
        ("util.vector", "size", "jvm.java.size", 0, 0),
        ("util.iterator", "next", "jvm.java.iterator_next", 0, 0),
        (
            "util.iterator",
            "hasnext",
            "jvm.java.iterator_has_next",
            0,
            0,
        ),
        ("util.listiterator", "next", "jvm.java.iterator_next", 0, 0),
        (
            "util.listiterator",
            "hasnext",
            "jvm.java.iterator_has_next",
            0,
            0,
        ),
        (
            "util.listiterator",
            "previous",
            "jvm.java.iterator_previous",
            0,
            0,
        ),
        (
            "util.listiterator",
            "hasprevious",
            "jvm.java.iterator_has_previous",
            0,
            0,
        ),
        (
            "util.listiterator",
            "nextindex",
            "jvm.java.iterator_next_index",
            0,
            0,
        ),
        (
            "util.listiterator",
            "previousindex",
            "jvm.java.iterator_previous_index",
            0,
            0,
        ),
        ("util.stack", "push", "jvm.java.add", 1, 1),
        ("util.stack", "pop", "collections.pop", 0, 0),
        ("util.stack", "peek", "jvm.java.peek_last", 0, 0),
        ("util.stack", "empty", "jvm.java.is_empty", 0, 0),
        ("util.stack", "size", "jvm.java.size", 0, 0),
        ("util.hashset", "add", "jvm.java.add", 1, 1),
        ("util.hashset", "size", "jvm.java.size", 0, 0),
        ("util.hashset", "remove", "jvm.java.list_remove", 1, 1),
        ("util.linkedhashset", "add", "jvm.java.add", 1, 1),
        ("util.linkedhashset", "size", "jvm.java.size", 0, 0),
        ("util.treeset", "add", "jvm.java.sorted_add", 1, 1),
        ("util.treeset", "first", "jvm.java.sorted_first", 0, 0),
        ("util.treeset", "last", "jvm.java.sorted_last", 0, 0),
        ("util.treeset", "higher", "jvm.java.sorted_higher", 1, 1),
        ("util.treeset", "lower", "jvm.java.sorted_lower", 1, 1),
        (
            "util.treeset",
            "descendingset",
            "jvm.java.sorted_descending_set",
            0,
            0,
        ),
        ("util.arraydeque", "addlast", "collections.push", 1, 1),
        ("util.arraydeque", "offer", "collections.push", 1, 1),
        ("util.arraydeque", "offerlast", "collections.push", 1, 1),
        ("util.arraydeque", "addfirst", "jvm.java.add_first", 1, 1),
        ("util.arraydeque", "offerfirst", "jvm.java.add_first", 1, 1),
        (
            "util.arraydeque",
            "removefirst",
            "jvm.java.remove_first",
            0,
            0,
        ),
        ("util.arraydeque", "removelast", "collections.pop", 0, 0),
        ("util.arraydeque", "peek", "jvm.java.peek_first", 0, 0),
        ("util.arraydeque", "peekfirst", "jvm.java.peek_first", 0, 0),
        ("util.arraydeque", "peeklast", "jvm.java.peek_last", 0, 0),
        ("util.arraydeque", "poll", "jvm.java.queue_poll", 0, 0),
        ("util.arraydeque", "push", "jvm.java.add_first", 1, 1),
        ("util.arraydeque", "pop", "jvm.java.poll_first", 0, 0),
        ("util.arraydeque", "size", "jvm.java.size", 0, 0),
        ("util.priorityqueue", "add", "jvm.java.priority_add", 1, 1),
        ("util.priorityqueue", "offer", "jvm.java.priority_add", 1, 1),
        ("util.priorityqueue", "peek", "jvm.java.priority_peek", 0, 0),
        ("util.priorityqueue", "poll", "jvm.java.queue_poll", 0, 0),
        ("util.priorityqueue", "size", "jvm.java.size", 0, 0),
        ("util.hashmap", "put", "jvm.java.map_put", 2, 2),
        ("util.hashmap", "get", "jvm.java.map_get", 1, 1),
        ("util.hashmap", "size", "jvm.java.map_size", 0, 0),
        ("util.hashmap", "isempty", "jvm.java.map_is_empty", 0, 0),
        ("util.hashmap", "clear", "jvm.java.map_clear", 0, 0),
        ("util.hashmap", "keyset", "jvm.java.map_key_set", 0, 0),
        ("util.hashmap", "values", "jvm.java.map_values", 0, 0),
        ("util.hashmap", "entryset", "jvm.java.entry_set", 0, 0),
        ("util.linkedhashmap", "put", "jvm.java.map_put", 2, 2),
        ("util.linkedhashmap", "get", "jvm.java.map_get", 1, 1),
        ("util.linkedhashmap", "size", "jvm.java.map_size", 0, 0),
        (
            "util.linkedhashmap",
            "isempty",
            "jvm.java.map_is_empty",
            0,
            0,
        ),
        ("util.linkedhashmap", "clear", "jvm.java.map_clear", 0, 0),
        ("util.linkedhashmap", "keyset", "jvm.java.map_key_set", 0, 0),
        ("util.linkedhashmap", "values", "jvm.java.map_values", 0, 0),
        ("util.linkedhashmap", "entryset", "jvm.java.entry_set", 0, 0),
        ("util.treemap", "put", "jvm.java.map_put", 2, 2),
        ("util.treemap", "get", "jvm.java.map_get", 1, 1),
        ("util.treemap", "size", "jvm.java.map_size", 0, 0),
        (
            "util.treemap",
            "keyset",
            "jvm.java.sorted_map_key_set",
            0,
            0,
        ),
        ("util.treemap", "values", "jvm.java.sorted_map_values", 0, 0),
        (
            "util.treemap",
            "firstkey",
            "jvm.java.sorted_map_first_key",
            0,
            0,
        ),
        (
            "util.treemap",
            "lastkey",
            "jvm.java.sorted_map_last_key",
            0,
            0,
        ),
        (
            "util.treemap",
            "higherkey",
            "jvm.java.sorted_map_higher_key",
            1,
            1,
        ),
        (
            "util.treemap",
            "lowerkey",
            "jvm.java.sorted_map_lower_key",
            1,
            1,
        ),
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

fn insert_java_io_methods(root: &mut Subtree) {
    const SPECS: &[(&str, &str, &str, u8, u8)] = &[
        ("io.bytearrayoutputstream", "size", "jvm.java.io_size", 0, 0),
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
        ("io.stringreader", "mark", "jvm.java.io_mark", 1, 1),
        ("io.stringreader", "reset", "jvm.java.io_reset_pos", 0, 0),
        ("io.stringreader", "skip", "jvm.java.io_skip", 1, 1),
        ("io.chararrayreader", "read", "jvm.java.io_read", 0, 1),
        ("io.chararrayreader", "mark", "jvm.java.io_mark", 1, 1),
        ("io.chararrayreader", "reset", "jvm.java.io_reset_pos", 0, 0),
        ("io.chararrayreader", "skip", "jvm.java.io_skip", 1, 1),
        ("io.inputstreamreader", "read", "jvm.java.io_read", 0, 1),
        ("io.inputstreamreader", "ready", "jvm.java.io_ready", 0, 0),
        ("io.bufferedinputstream", "read", "jvm.java.io_read", 0, 1),
        ("io.filterinputstream", "read", "jvm.java.io_read", 0, 1),
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
    ] {
        merge_type_member_returns(root, type_path, returns);
    }
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
        ("util.collections", "singleton", "jvm.java.list_of"),
        ("util.collections", "singletonmap", "jvm.java.map_of"),
        ("util.collections", "singletonlist", "jvm.java.list_of"),
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
        ("time.yearmonth", "parse", "jvm.java.local_date_parse"),
        ("time.monthday", "parse", "jvm.java.local_date_parse"),
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
            ][..],
        ),
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
        ("get", "jvm.java.identity", 1, 1),
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
        ("tostring", "jvm.java.time_to_string", 0, 0),
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
        ("with", "jvm.java.identity", 2, 2),
        ("isbefore", "jvm.java.instant_is_before", 1, 1),
        ("isafter", "jvm.java.instant_is_after", 1, 1),
        ("compareto", "jvm.java.instant_compare_to", 1, 1),
        ("tostring", "jvm.java.time_to_string", 0, 0),
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
                ("toLocalDate", "java.time.LocalDate"),
                ("toLocalTime", "java.time.LocalTime"),
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
                field_gui: Vec::new(),
                value_equality: false }),
            ctor_call: Some(Box::new(common_emit("jvm.java.net.url_new"))),
            statics: Subtree::new(),
            methods: url_methods,
            member_returns: url_returns },
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
            member_returns: Default::default() },
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
            member_returns: Default::default() },
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
                field_gui: Vec::new(),
                value_equality: false }),
            ctor_call: Some(Box::new(common_emit("jvm.java.net.uri_new"))),
            statics: uri_statics,
            methods: uri_methods,
            member_returns: uri_returns },
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
    pub intrinsic: Option<&'static str> }

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
        intrinsic }
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
        _ => return None })
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
                        field_gui: Vec::new(),
                        value_equality: false }),
                    ctor_call: Some(Box::new(ctor_call)),
                    statics: Subtree::new(),
                    methods: Subtree::new(),
                    member_returns: Default::default() },
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
        insert_java_util_collection_statics(&mut root);
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
}
