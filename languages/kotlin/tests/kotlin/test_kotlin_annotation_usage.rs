kotlin_run_test!(
    test_annotation_with_enum_and_boolean_args,
    r#"
        enum class Level { OFF, WARN, ON }

        @Target(AnnotationTarget.CLASS)
        annotation class Flag(val level: Level, val active: Boolean = true)

        @Flag(level = Level.ON, active = true)
        class Processor

        fun main() {
            println(Processor::class.simpleName)
        }
    "#,
    &["Processor"]
);

kotlin_run_test!(
    test_annotation_with_kclass_argument,
    r#"
        @Target(AnnotationTarget.CLASS)
        annotation class DelegateType(val impl: kotlin.reflect.KClass<*>)

        @DelegateType(impl = String::class)
        class Storage

        fun main() {
            println(Storage::class.simpleName)
        }
    "#,
    &["Storage"]
);

kotlin_run_test!(
    test_repeatable_annotations_are_allowed_on_target,
    r#"
        @Target(AnnotationTarget.CLASS)
        @Repeatable
        annotation class Tag(val name: String)

        @Tag("a")
        @Tag("b")
        class Dual

        fun main() {
            println(Dual::class.simpleName)
        }
    "#,
    &["Dual"]
);

kotlin_run_test!(
    test_nested_annotation_arguments,
    r#"
        enum class Kind { ALPHA, BETA }
        annotation class Meta(val kind: Kind)
        annotation class Bundle(val metas: Array<Meta>)

        @Bundle([Meta(Kind.ALPHA), Meta(Kind.BETA)])
        class Combined

        fun main() {
            println(Combined::class.simpleName)
        }
    "#,
    &["Combined"]
);

kotlin_run_test!(
    test_annotation_with_long_and_char_defaults,
    r#"
        @Target(AnnotationTarget.FUNCTION)
        annotation class Config(val symbol: Char = 'x', val limit: Long = 10)

        @Config
        fun value() = 12

        fun main() {
            println(value())
        }
    "#,
    &["12"]
);

kotlin_run_test!(
    test_file_targeted_annotations_do_not_change_runtime_output,
    r#"
        @Target(AnnotationTarget.FILE)
        annotation class FileMeta

        @FileMeta
        fun marker() = true

        fun main() {
            println(marker())
        }
    "#,
    &["true"]
);

kotlin_run_test!(
    test_constructor_annotation_in_generic_class,
    r#"
        @Target(AnnotationTarget.CONSTRUCTOR)
        annotation class Build

        class Holder<@Build T>(val value: T)

        fun main() {
            println(Holder(3).value)
        }
    "#,
    &["3"]
);

kotlin_run_test!(
    test_type_alias_with_annotated_type_parameter,
    r#"
        @Target(AnnotationTarget.TYPE)
        annotation class Tainted

        typealias TaggedInt = @Tainted Int

        fun main() {
            val v: TaggedInt = 9
            println(v)
        }
    "#,
    &["9"]
);

kotlin_run_test!(
    test_receiver_annotation_syntax_compiles,
    r##"
        @Target(AnnotationTarget.RECEIVER)
        annotation class Ext

        class Token

        @Ext
        fun Token.wrap(prefix: String) = prefix + "#"

        fun main() {
            println(Token().wrap("x"))
        }
    "##,
    &["x#"]
);

kotlin_run_test!(
    test_parameterized_local_annotation_is_parsed,
    r#"
        @Target(AnnotationTarget.VALUE_PARAMETER)
        annotation class Marker(val id: Int)

        fun compute(@Marker(42) value: Int): Int = value + 1

        fun main() {
            println(compute(4))
        }
    "#,
    &["5"]
);
