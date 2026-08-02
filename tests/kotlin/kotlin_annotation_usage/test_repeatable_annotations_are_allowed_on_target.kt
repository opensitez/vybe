// vybe-test: kotlin/kotlin_annotation_usage/test_repeatable_annotations_are_allowed_on_target
// origin: languages/kotlin/tests/kotlin/test_kotlin_annotation_usage.rs

@Target(AnnotationTarget.CLASS)
        @Repeatable
        annotation class Tag(val name: String)

        @Tag("a")
        @Tag("b")
        class Dual

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Dual::class.simpleName).toString(), "Dual")
        }
