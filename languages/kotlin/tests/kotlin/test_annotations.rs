use crate::helpers::run_prints;

#[test]
fn test_annotation_parsing() {
    let out = run_prints(
        r#"
        @Deprecated
        fun oldFunction() {
            println("deprecated function executed")
        }

        fun main() {
            oldFunction()
        }
    "#,
    );
    assert_eq!(out, &["deprecated function executed"]);
}

#[test]
fn test_annotation_with_arguments() {
    let out = run_prints(
        r#"
        @Suppress("UNCHECKED_CAST")
        fun castFunc() {
            println("annotated with args")
        }

        fun main() {
            castFunc()
        }
    "#,
    );
    assert_eq!(out, &["annotated with args"]);
}

#[test]
fn test_annotation_on_class_and_member() {
    let out = run_prints(
        r#"
        @Deprecated
        class Legacy {
            fun name(): String = "legacy"
        }

        @Suppress("UNUSED_PARAMETER")
        fun tagged(@Deprecated code: Int): String {
            return "tagged"
        }

        fun main() {
            val legacy = Legacy()
            println(legacy.name())
            println(tagged(1))
        }
    "#,
    );
    assert_eq!(out, &["legacy", "tagged"]);
}

#[test]
fn test_multiple_annotations() {
    let out = run_prints(
        r#"
        @Deprecated
        @Suppress("UNUSED_VARIABLE")
        fun deprecated_function() {
            println("deprecated_function")
        }

        fun main() {
            deprecated_function()
        }
    "#,
    );
    assert_eq!(out, &["deprecated_function"]);
}

#[test]
fn test_annotation_on_property() {
    let out = run_prints(
        r#"
        class Versioned {
            @Deprecated("legacy field")
            val tag = "v1"
        }

        fun main() {
            val v = Versioned()
            println(v.tag)
        }
    "#,
    );
    assert_eq!(out, &["v1"]);
}

#[test]
fn test_annotation_with_named_args() {
    let out = run_prints(
        r#"
        @Suppress("UNUSED_PARAMETER")
        fun log(@Deprecated code: Int): String {
            return code.toString()
        }

        fun main() {
            println(log(7))
        }
    "#,
    );
    assert_eq!(out, &["7"]);
}

#[test]
fn test_annotation_on_interface() {
    let out = run_prints(
        r#"
        @Deprecated("old interface")
        interface Marker {
            fun label(): String
        }

        class Tag : Marker {
            override fun label(): String {
                return "tagged"
            }
        }

        fun main() {
            val m: Marker = Tag()
            println(m.label())
        }
    "#,
    );
    assert_eq!(out, &["tagged"]);
}

#[test]
fn test_annotation_on_object_declaration() {
    let out = run_prints(
        r#"
        @Deprecated("legacy")
        class Notifier {
            companion object {
                fun ping(): String {
                    return "pong"
                }
            }
        }

        fun main() {
            println(Notifier.ping())
        }
    "#,
    );
    assert_eq!(out, &["pong"]);
}

#[test]
fn test_annotation_multi_line_stack() {
    let out = run_prints(
        r#"
        @Deprecated("legacy")
        @Suppress("UNUSED_VARIABLE")
        fun taggedFunction() {
            println("stack")
        }

        fun main() {
            taggedFunction()
        }
    "#,
    );
    assert_eq!(out, &["stack"]);
}

#[test]
fn test_annotation_on_constructor_parameter() {
    let out = run_prints(
        r#"
        class User(
            @Deprecated("unused") val id: Int,
            @Suppress("UNUSED_PARAMETER") val name: String
        )

        fun main() {
            val user = User(1, "alice")
            println(user.id)
            println(user.name)
        }
    "#,
    );
    assert_eq!(out, &["1", "alice"]);
}

#[test]
fn test_annotation_with_multiple_arguments() {
    let out = run_prints(
        r#"
        @Suppress("UNUSED_VARIABLE", "NAME_SHADOWING")
        fun annotated() {
            println("multi")
        }

        fun main() {
            annotated()
        }
    "#,
    );
    assert_eq!(out, &["multi"]);
}

#[test]
fn test_annotation_with_typealias() {
    let out = run_prints(
        r#"
        @Deprecated("old alias")
        class Greeting {
            val message: String = "hi"
        }

        fun main() {
            val msg = Greeting()
            println(msg.message)
        }
    "#,
    );
    assert_eq!(out, &["hi"]);
}

#[test]
fn test_annotation_extension_receiver() {
    let out = run_prints(
        r#"
        @Deprecated("old")
        fun String.highlight(): String {
            return "<<" + this + ">>"
        }

        fun main() {
            println("ok".highlight())
        }
    "#,
    );
    assert_eq!(out, &["<<ok>>"]);
}

#[test]
fn test_annotation_on_local_variable() {
    let out = run_prints(
        r#"
        fun main() {
            @Deprecated("temp")
            val status = "pending"
            println(status)
        }
    "#,
    );
    assert_eq!(out, &["pending"]);
}

#[test]
fn test_annotation_on_function_overload_target() {
    let out = run_prints(
        r#"
        @Suppress("UNUSED_PARAMETER")
        fun label(a: Int) {
            println(a)
        }

        fun main() {
            label(9)
        }
    "#,
    );
    assert_eq!(out, &["9"]);
}

#[test]
fn test_annotation_order_is_irrelevant_to_execution() {
    let out = run_prints(
        r#"
        @Suppress("UNUSED")
        @Deprecated("deprecated")
        fun marker(value: Int): Int {
            return value + 1
        }

        fun main() {
            println(marker(4))
        }
    "#,
    );
    assert_eq!(out, &["5"]);
}

#[test]
fn test_annotation_local_function_and_call() {
    let out = run_prints(
        r#"
fun main() { @Deprecated("local") fun local() { println("local") }; local() }
"#,
    );
    assert_eq!(out, &["local"]);
}

#[test]
fn test_annotation_on_top_level_parameter() {
    let out = run_prints(
        r#"
@Suppress("UNUSED_PARAMETER") fun marker(@Deprecated("x") x: Int): Int { return x + 2 }; fun main() { println(marker(4)) }
"#,
    );
    assert_eq!(out, &["6"]);
}

#[test]
fn test_annotation_on_class_constructor_parameter() {
    let out = run_prints(
        r#"
class Packet(@Deprecated("old") val id: Int) ; fun main() { val p = Packet(9); println(p.id) }
"#,
    );
    assert_eq!(out, &["9"]);
}

#[test]
fn test_annotation_on_object_member() {
    let out = run_prints(
        r#"
class Sender {
    companion object {
        @Deprecated("legacy")
        @Suppress("UNUSED")
        val token = "go"

        @Suppress("UNUSED")
        fun code(): String = token
    }
}

fun main() {
    println(Sender.code())
}
"#,
    );
    assert_eq!(out, &["go"]);
}

#[test]
fn test_annotation_with_class_and_method() {
    let out = run_prints(
        r#"
@Deprecated("old") class Marker { @Suppress("UNUSED_PARAMETER") fun tagged(@Deprecated("id") id: Int): Int { return id } }; fun main() { println(Marker().tagged(11)) }
"#,
    );
    assert_eq!(out, &["11"]);
}

#[test]
fn test_annotation_on_secondary_constructor() {
    let out = run_prints(
        r#"
class Session {
    val id: Int

    constructor() {
        this.id = 1
    }

    @Deprecated("secondary")
    constructor(id: Int) {
        this.id = id
    }
}

fun main() {
    val s = Session(5)
    println(s.id)
}
"#,
    );
    assert_eq!(out, &["5"]);
}

#[test]
fn test_annotation_companion_function() {
    let out = run_prints(
        r#"
class Factory {
            companion object {
                @Deprecated("legacy") fun create(): Int = 21
            }
        }
        fun main() { println(Factory.create()) }
"#,
    );
    assert_eq!(out, &["21"]);
}

#[test]
fn test_annotation_on_function_expression() {
    let out = run_prints(
        r#"
@Suppress("UNUSED") fun action(): String { return "done" }; fun main() { println(action()) }
"#,
    );
    assert_eq!(out, &["done"]);
}

#[test]
fn test_annotation_on_extension_receiver() {
    let out = run_prints(
        r#"
@Deprecated("legacy") fun String.wrap(): String = "[" + this + "]"; fun main() { println("ok".wrap()) }
"#,
    );
    assert_eq!(out, &["[ok]"]);
}

#[test]
fn test_annotation_multi_stack_local() {
    let out = run_prints(
        r#"
fun main() { @Suppress("UNUSED_PARAMETER") @Deprecated("old") fun local(v: Int): Int { return v + 1 }; println(local(6)) }
"#,
    );
    assert_eq!(out, &["7"]);
}

#[test]
fn test_annotation_on_typealias_value() {
    let out = run_prints(
        r#"
@Suppress("UNUSED")
@Deprecated("legacy")
class Score(val value: Int = 13)

fun main() {
    val value = Score(13)
    println(value.value)
}
"#,
    );
    assert_eq!(out, &["13"]);
}

#[test]
fn test_annotation_on_interface_member() {
    let out = run_prints(
        r#"
interface Logger { @Deprecated("old") fun emit(): Int }
        class Console : Logger { override fun emit(): Int = 3 }
        fun main() { val l: Logger = Console(); println(l.emit()) }
"#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_annotation_on_value_member() {
    let out = run_prints(
        r#"
class Counter {
            @Deprecated("counter")
            val total = 4
        }
        fun main() { println(Counter().total) }
"#,
    );
    assert_eq!(out, &["4"]);
}

#[test]
fn test_annotation_with_custom_message() {
    let out = run_prints(
        r#"
@Suppress("UNUSED_PARAMETER") fun tagged(@Suppress("UNUSED") x: Int, @Deprecated("bad") y: Int): Int { return x + y }; fun main() { println(tagged(2, 3)) }
"#,
    );
    assert_eq!(out, &["5"]);
}

#[test]
fn test_annotation_declaration_and_usage_with_parameters() {
    let out = run_prints(
        r#"
        annotation class Tag(val label: String)

        @Tag("service")
        fun service() {
            println("service")
        }

        fun main() {
            service()
        }
    "#,
    );
    assert_eq!(out, &["service"]);
}

#[test]
fn test_annotation_with_defaulted_constructor_argument() {
    let out = run_prints(
        r#"
        annotation class Flag(val level: String = "low")

        @Flag
        @Flag("high")
        fun run(level: String = "x"): String {
            return level
        }

        fun main() {
            println(run())
        }
    "#,
    );
    assert_eq!(out, &["x"]);
}

#[test]
fn test_annotation_meta_targets() {
    let out = run_prints(
        r#"
        @Target(AnnotationTarget.CLASS)
        @Retention(AnnotationRetention.RUNTIME)
        annotation class TypeTag

        @TypeTag
        class Marked

        fun main() {
            val item: Any = Marked()
            println(item::class.simpleName)
        }
    "#,
    );
    assert_eq!(out, &["Marked"]);
}

#[test]
fn test_annotation_array_arguments() {
    let out = run_prints(
        r#"
        annotation class Labels(vararg val values: String)

        @Labels("a", "b", "c")
        fun report(): String {
            return "counted"
        }

        fun main() {
            println(report())
        }
    "#,
    );
    assert_eq!(out, &["counted"]);
}

#[test]
fn test_multiple_custom_annotations_on_member() {
    let out = run_prints(
        r#"
        annotation class One
        annotation class Two

        @One
        @Two
        fun both(): Int {
            return 2 + 3
        }

        fun main() {
            println(both())
        }
    "#,
    );
    assert_eq!(out, &["5"]);
}

#[test]
fn test_annotation_usage_on_local_type_alias_member() {
    let out = run_prints(
        r#"
        annotation class Local

        fun main() {
            @Local
            class Holder(val value: Int)
            val h = Holder(9)
            println(h.value)
        }
    "#,
    );
    assert_eq!(out, &["9"]);
}

#[test]
fn test_enum_entry_annotation_support() {
    let out = run_prints(
        r#"
        annotation class State

        enum class Mode {
            @State
            OFF,

            @State
            ON
        }

        fun main() {
            println(Mode.ON.name)
        }
    "#,
    );
    assert_eq!(out, &["ON"]);
}

#[test]
fn test_annotation_class_target_specification() {
    let out = run_prints(
        r#"
        @Target(AnnotationTarget.CLASS, AnnotationTarget.FUNCTION)
        @Retention(AnnotationRetention.RUNTIME)
        annotation class Role(val name: String)

        @Role("service")
        class Service {
            @Role("entry")
            fun start(): String = "ready"
        }

        fun main() {
            val service = Service()
            println(service.start())
        }
    "#,
    );
    assert_eq!(out, &["ready"]);
}

#[test]
fn test_file_annotation_target_is_parsed() {
    let out = run_prints(
        r#"
        @file:Suppress("UNUSED_VARIABLE")
        fun ignored() {}

        fun main() {
            println("ok")
        }
    "#,
    );
    assert_eq!(out, &["ok"]);
}

#[test]
fn test_getter_use_site_target_annotation() {
    let out = run_prints(
        r#"
        class Record {
            @get:Deprecated("legacy")
            val label: String = "label"
        }

        fun main() {
            val r = Record()
            println(r.label)
        }
    "#,
    );
    assert_eq!(out, &["label"]);
}

#[test]
fn test_setter_use_site_target_annotation() {
    let out = run_prints(
        r#"
        class Holder {
            @set:Suppress("UNUSED")
            var value: Int = 0
        }

        fun main() {
            val h = Holder()
            h.value = 6
            println(h.value)
        }
    "#,
    );
    assert_eq!(out, &["6"]);
}

#[test]
fn test_annotation_with_enum_argument() {
    let out = run_prints(
        r#"
        enum class Level { LOW, MEDIUM, HIGH }

        annotation class Tier(val level: Level)

        @Tier(Level.HIGH)
        fun service() = "tiered"

        fun main() {
            println(service())
        }
    "#,
    );
    assert_eq!(out, &["tiered"]);
}

#[test]
fn test_annotation_with_nested_custom_argument() {
    let out = run_prints(
        r#"
        annotation class Tag(val value: String)
        annotation class Marker(val tag: Tag)

        @Marker(Tag("critical"))
        class Item

        fun main() {
            println(Item().javaClass.simpleName)
        }
    "#,
    );
    assert_eq!(out, &["Item"]);
}
