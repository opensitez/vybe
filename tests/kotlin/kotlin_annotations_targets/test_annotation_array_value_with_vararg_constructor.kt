// vybe-test: kotlin/kotlin_annotations_targets/test_annotation_array_value_with_vararg_constructor
// origin: languages/kotlin/tests/kotlin/test_kotlin_annotations_targets.rs

@Target(AnnotationTarget.CLASS)
        annotation class Tags(val items: Array<String>)

        @Tags(["alpha", "beta"])
        class Target

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Target::class.simpleName).toString(), "Target")
        }
