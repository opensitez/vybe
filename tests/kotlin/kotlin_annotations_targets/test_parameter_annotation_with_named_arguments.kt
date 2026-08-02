// vybe-test: kotlin/kotlin_annotations_targets/test_parameter_annotation_with_named_arguments
// origin: languages/kotlin/tests/kotlin/test_kotlin_annotations_targets.rs

@Target(AnnotationTarget.VALUE_PARAMETER)
        annotation class NameTag(val value: String)

        fun greet(@NameTag("primary") who: String): String {
            return "hi $who"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((greet("Ada")).toString(), "hi Ada")
        }
