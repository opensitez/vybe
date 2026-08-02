// vybe-test: kotlin/kotlin_annotations_targets/test_expression_body_uses_annotated_function_and_prints_result
// origin: languages/kotlin/tests/kotlin/test_kotlin_annotations_targets.rs

@Target(AnnotationTarget.FUNCTION)
        annotation class InlineLike

        @InlineLike
        fun add(a: Int, b: Int): Int = a + b

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((add(2, 3)).toString(), "5")
        }
