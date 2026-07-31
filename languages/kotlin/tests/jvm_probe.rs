//! Cross-language JDK resolution: Kotlin reaching `java.*` through the COMMON
//! RESOLVER, with ZERO `java.*` declarations in Kotlin's own profile.
//!
//! This is the load-bearing claim of javakotlinmigration.md, so it is a test
//! and not a scratch file: `platforms/jvm` owns the declarations, Kotlin
//! declares only tree data (`type_scopes` + `kind = "tree-ambient"`), and the
//! resolver does the rest — the same way csharp/vb reach `dotnet.*`.

fn register_both() {
    static ONCE: std::sync::Once = std::sync::Once::new();
    ONCE.call_once(vybe_language_kotlin::register);
}

fn try_compile(src: &str) -> Result<Vec<vybe_runtime::Chunk>, String> {
    register_both();
    let module = vybe_language_kotlin::parse(src)?;
    let profile = vybe_compiler::profile::parse_profile(vybe_language_kotlin::profile_source())
        .map_err(|e| format!("profile parse failed: {}", e))?;
    vybe_compiler::primitives::Compiler::with_profile(profile).compile(&module)
}

fn probe(label: &str, src: &str) {
    use std::sync::{Arc, Mutex};
    use vybe_runtime::{HostContext, VM, Value};
    let chunks = match try_compile(src) {
        Ok(c) => c,
        Err(e) => {
            println!("[{label}] COMPILE FAIL: {e}");
            return;
        }
    };
    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let out = output.clone();
    vybe_compiler::primitives::platforms::init_platforms(&mut vm);
    vm.register_host_fn(
        "wasi:logging/logging",
        "log",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let s: Vec<String> = args.iter().map(|a| format!("{}", a)).collect();
            out.lock().unwrap().push(s.join(" "));
            Value::Null
        }),
    );
    vybe_compiler::primitives::platforms::finalize_platforms(&mut vm);
    match vm.run(chunks) {
        Ok(_) => println!("[{label}] RAN → {:?}", output.lock().unwrap()),
        Err(e) => println!("[{label}] RUNTIME TRAP: {e}"),
    }
}

#[test]
fn tree_shape() {
    register_both();
    // Force the lazy platform/language tree registration the resolver does.
    let _ = try_compile("fun main() { println(1) }\nmain()\n");
    for path in [
        "java",
        "java.time",
        "java.time.instant",
        "java.time.instant.parse",
        "java.java",
        "java.java.time.instant.parse",
        "java.util.arraylist",
        "java.java.util.arraylist",
        "java.lang.math.max",
        "java.util.arrays",
        "java.util.arrays.aslist",
        "java.util.arrays.sort",
        "java.time.instant.ofepochsecond",
        "java.time.instant.now",
        "java.util.uuid.randomuuid",
    ] {
        let segs: Vec<&str> = path.split('.').collect();
        println!(
            "  tree[{path}] = {:?}",
            vybe_compiler::primitives::namespaces::resolve_path(&segs)
        );
    }
}

#[test]
fn jvm_probe() {
    // ADVISOR ITEM 1: a lowercase probe against a `java.`-PREFIXED emit target.
    // `objects.equals` routes to the SHARED `common:object.equals` and never
    // enters Java's emitter, so it does not test the plan's real claim —
    // that `common:java.*` reaches Java's 21k-line emitter via
    // `languages::emit_dispatch_for("java")`.
    probe(
        "fq-lowercase-instant-parse",
        r#"fun main() {
    val t = java.time.instant.parse("2020-01-01T00:00:00Z")
    println("instant ok")
}
main()
"#,
    );
    // Case hypothesis: tree keys are lowercase; `canon()` preserves case for
    // case-sensitive languages. An all-lowercase spelling of the same path
    // should therefore resolve where the real Java spelling does not.
    probe(
        "fq-lowercase-objects-equals",
        r#"fun main() {
    val b = java.util.objects.equals(1, 1)
    if (b) { println("EQ") } else { println("NE") }
}
main()
"#,
    );
    // Baselines: does the probe harness itself work at all?
    probe("baseline-println", "fun main() { println(1) }\nmain()\n");
    // `java.*` value produced but NOT printed — separates the call from print.
    probe(
        "fq-objects-equals-nonprint",
        r#"fun main() {
    val b = java.util.Objects.equals(1, 1)
    if (b) { println("EQ") } else { println("NE") }
}
main()
"#,
    );
    probe(
        "fq-instant-parse-nonprint",
        r#"fun main() {
    val t = java.time.Instant.parse("2020-01-01T00:00:00Z")
    println("got instant")
}
main()
"#,
    );
    probe(
        "fq-instant-parse",
        r#"fun main() {
    val t = java.time.Instant.parse("2020-01-01T00:00:00Z")
    println(t)
}
main()
"#,
    );
    probe(
        "import-instant",
        r#"import java.time.Instant
fun main() {
    val t = Instant.parse("2020-01-01T00:00:00Z")
    println(t)
}
main()
"#,
    );
    probe(
        "fq-arraylist-new",
        r#"fun main() {
    val xs = java.util.ArrayList<Int>()
    println(xs)
}
main()
"#,
    );
    probe(
        "import-arraylist-new",
        r#"import java.util.ArrayList
fun main() {
    val xs = ArrayList<Int>()
    println(xs)
}
main()
"#,
    );
    probe(
        "fq-objects-equals",
        r#"fun main() {
    println(java.util.Objects.equals(1, 1))
}
main()
"#,
    );
    probe(
        "import-file",
        r#"import java.io.File
fun main() {
    val f = File("x")
    println(f)
}
main()
"#,
    );
}

/// ADVISOR ITEM 3: enumerate EVERY `java.*` profile key through the tree and
/// report the misses. `insert_path` drops silently on a leaf/namespace
/// collision, so a dropped leaf is otherwise invisible.
#[test]
fn all_java_keys_resolve() {
    register_both();
    let _ = try_compile("fun main() { println(1) }\nmain()\n");
    let src = vybe_language_java::profile_source();
    let mut keys: Vec<String> = Vec::new();
    for line in src.lines() {
        let line = line.trim();
        let Some(rest) = line.strip_prefix('"') else {
            continue;
        };
        let Some(end) = rest.find('"') else { continue };
        let key = &rest[..end];
        if key.starts_with("java.") && rest[end + 1..].trim_start().starts_with('=') {
            keys.push(key.to_string());
        }
    }
    keys.sort();
    keys.dedup();
    let mut missed = Vec::new();
    for key in &keys {
        let lower = key.to_lowercase();
        let segs: Vec<&str> = lower.split('.').collect();
        if vybe_compiler::primitives::namespaces::resolve_path(&segs).is_none() {
            missed.push(key.clone());
        }
    }
    println!("java.* profile keys: {}", keys.len());
    println!("unreachable in tree: {}", missed.len());
    for m in &missed {
        println!("  MISS {m}");
    }
}

#[test]
fn ast_shape() {
    register_both();
    let src = "import java.util.ArrayList\nfun main() {\n    val a = java.util.ArrayList<Int>()\n    val b = ArrayList<Int>()\n}\n";
    match vybe_language_kotlin::parse(src) {
        Ok(m) => println!("AST: {:#?}", m.body),
        Err(e) => println!("PARSE FAIL: {e}"),
    }
}

/// Kotlin's OWN names must not be hijacked by the `java.lang` ambient. The
/// ambient is a pure fallback, but `Double`, `Thread` and `StringBuilder` all
/// exist under `java.lang` in the tree, so this is the collision to watch.
#[test]
fn kotlin_names_not_shadowed_by_java_ambients() {
    probe("kt-string-ops", r#"fun main() {
    val s = "hello"
    println(s.length)
    println(s.uppercase())
    println(s.substring(1, 3))
}
main()
"#);
    probe("kt-double", r#"fun main() {
    val d: Double = 2.5
    println(d + 1.0)
}
main()
"#);
    probe("kt-stringbuilder", r#"fun main() {
    val sb = StringBuilder()
    sb.append("a")
    sb.append("b")
    println(sb.toString())
}
main()
"#);
    probe("kt-listof", r#"fun main() {
    val xs = listOf(1, 2, 3)
    println(xs.size)
}
main()
"#);
}

/// Breadth of the JDK surface reachable FROM KOTLIN, with zero `java.*`
/// declarations in Kotlin's profile. Each line is a different package.
#[test]
fn kotlin_reaches_the_jdk_surface() {
    probe("jdk-uuid", r#"fun main() {
    val u = java.util.UUID.randomUUID()
    println(u != null)
}
main()
"#);
    probe("jdk-bigint", r#"fun main() {
    val b = java.math.BigInteger("123")
    println(b)
}
main()
"#);
    probe("jdk-hashmap", r#"import java.util.HashMap
fun main() {
    val m = HashMap<String, Int>()
    println(m)
}
main()
"#);
    probe("jdk-duration", r#"fun main() {
    val d = java.time.Duration.ofSeconds(90)
    println(d != null)
}
main()
"#);
    probe("jdk-zoneid", r#"fun main() {
    val z = java.time.ZoneId.of("UTC")
    println(z)
}
main()
"#);
    probe("jdk-objects-hash", r#"fun main() {
    println(java.util.Objects.isNull(null))
}
main()
"#);
}

/// Statics resolve; do INSTANCE methods on the returned JDK object?
#[test]
fn kotlin_jdk_instance_members() {
    probe("static-only", r#"fun main() {
    val d = java.time.LocalDate.parse("2024-06-15")
    println("parsed")
}
main()
"#);
    probe("instance-prop", r#"fun main() {
    val d = java.time.LocalDate.parse("2024-06-15")
    println(d.year)
}
main()
"#);
    probe("instance-method", r#"fun main() {
    val d = java.time.LocalDate.parse("2024-06-15")
    println(d.isLeapYear())
}
main()
"#);
    probe("list-instance-method", r#"import java.util.ArrayList
fun main() {
    val xs = ArrayList<Int>()
    xs.add(1)
    println(xs.size)
}
main()
"#);
}
