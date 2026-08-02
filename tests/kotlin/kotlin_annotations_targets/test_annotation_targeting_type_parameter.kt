// vybe-test: kotlin/kotlin_annotations_targets/test_annotation_targeting_type_parameter
// origin: languages/kotlin/tests/kotlin/test_kotlin_annotations_targets.rs

@Target(AnnotationTarget.TYPE_PARAMETER)
        annotation class TypeOnly

        class Wrapper<@TypeOnly T>(val value: T)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Wrapper(7).value).toString(), "7")
        }
