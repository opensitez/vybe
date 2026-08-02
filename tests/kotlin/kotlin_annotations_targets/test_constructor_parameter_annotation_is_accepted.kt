// vybe-test: kotlin/kotlin_annotations_targets/test_constructor_parameter_annotation_is_accepted
// origin: languages/kotlin/tests/kotlin/test_kotlin_annotations_targets.rs

@Target(AnnotationTarget.VALUE_PARAMETER)
        annotation class Required

        class Service(@Required val host: String) {
            fun describe() = host
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Service("edge").describe()).toString(), "edge")
        }
