kotlin_run_test!(
    test_file_level_class_annotation_compiles_to_program_output,
    r#"
        @Target(AnnotationTarget.CLASS)
        annotation class Marker

        @Marker
        class Service

        fun main() {
            val service = Service()
            println(service::class.simpleName)
        }
    "#,
    &[{"Service"}]
);

kotlin_run_test!(
    test_annotation_parameterized_targeted_at_function,
    r#"
        @Target(AnnotationTarget.FUNCTION)
        annotation class Route(val path: String = "/")

        @Route(path = "/ping")
        fun ping() = "pong"

        fun main() {
            println(ping())
        }
    "#,
    &["pong"]
);

kotlin_run_test!(
    test_property_and_getter_level_annotations,
    r#"
        @Target(AnnotationTarget.PROPERTY, AnnotationTarget.FIELD)
        annotation class Visible

        class Box {
            @Visible
            var value: Int = 7
        }

        fun main() {
            val b = Box()
            println(b.value)
        }
    "#,
    &["7"]
);

kotlin_run_test!(
    test_parameter_annotation_with_named_arguments,
    r#"
        @Target(AnnotationTarget.VALUE_PARAMETER)
        annotation class NameTag(val value: String)

        fun greet(@NameTag("primary") who: String): String {
            return "hi $who"
        }

        fun main() {
            println(greet("Ada"))
        }
    "#,
    &["hi Ada"]
);

kotlin_run_test!(
    test_multiple_compatible_annotations_on_same_target,
    r#"
        @Target(AnnotationTarget.CLASS)
        annotation class A(val value: String)

        @Target(AnnotationTarget.CLASS)
        annotation class B(val value: Int)

        @A("layer")
        @B(3)
        class Tagged

        fun main() {
            println(Tagged::class.simpleName)
        }
    "#,
    &["Tagged"]
);

kotlin_run_test!(
    test_constructor_parameter_annotation_is_accepted,
    r#"
        @Target(AnnotationTarget.VALUE_PARAMETER)
        annotation class Required

        class Service(@Required val host: String) {
            fun describe() = host
        }

        fun main() {
            println(Service("edge").describe())
        }
    "#,
    &["edge"]
);

kotlin_run_test!(
    test_expression_body_uses_annotated_function_and_prints_result,
    r#"
        @Target(AnnotationTarget.FUNCTION)
        annotation class InlineLike

        @InlineLike
        fun add(a: Int, b: Int): Int = a + b

        fun main() {
            println(add(2, 3))
        }
    "#,
    &["5"]
);

kotlin_run_test!(
    test_annotation_array_value_with_vararg_constructor,
    r#"
        @Target(AnnotationTarget.CLASS)
        annotation class Tags(val items: Array<String>)

        @Tags(["alpha", "beta"])
        class Target

        fun main() {
            println(Target::class.simpleName)
        }
    "#,
    &["Target"]
);

kotlin_run_test!(
    test_receiver_annotation_compiles_on_extension_receiver,
    r#"
        @Target(AnnotationTarget.RECEIVER)
        annotation class ReceiverMarker

        class Box {
            fun text() = "ok"
        }

        @ReceiverMarker
        fun Box.announce(): String = this.text()

        fun main() {
            println(Box().announce())
        }
    "#,
    &["ok"]
);

kotlin_run_test!(
    test_annotation_targeting_type_parameter,
    r#"
        @Target(AnnotationTarget.TYPE_PARAMETER)
        annotation class TypeOnly

        class Wrapper<@TypeOnly T>(val value: T)

        fun main() {
            println(Wrapper(7).value)
        }
    "#,
    &["7"]
);

kotlin_run_test!(
    test_annotation_targeting_field_and_local_parameter_combo,
    r#"
        @Target(AnnotationTarget.FIELD)
        annotation class FieldMark

        @Target(AnnotationTarget.VALUE_PARAMETER)
        annotation class FieldArg

        class Payload(@FieldArg val marker: String) {
            @FieldMark
            val copyMarker: String = marker
        }

        fun main() {
            println(Payload("x").copyMarker)
        }
    "#,
    &["x"]
);
