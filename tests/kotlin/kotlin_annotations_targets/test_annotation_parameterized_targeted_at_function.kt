// vybe-test: kotlin/kotlin_annotations_targets/test_annotation_parameterized_targeted_at_function
// origin: languages/kotlin/tests/kotlin/test_kotlin_annotations_targets.rs

@Target(AnnotationTarget.FUNCTION)
        annotation class Route(val path: String = "/")

        @Route(path = "/ping")
        fun ping() = "pong"

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((ping()).toString(), "pong")
        }
