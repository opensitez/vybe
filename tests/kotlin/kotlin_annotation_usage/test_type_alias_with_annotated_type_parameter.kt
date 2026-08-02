// vybe-test: kotlin/kotlin_annotation_usage/test_type_alias_with_annotated_type_parameter
// origin: languages/kotlin/tests/kotlin/test_kotlin_annotation_usage.rs

@Target(AnnotationTarget.TYPE)
        annotation class Tainted

        typealias TaggedInt = @Tainted Int

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val v: TaggedInt = 9
            __check((v).toString(), "9")
        }
