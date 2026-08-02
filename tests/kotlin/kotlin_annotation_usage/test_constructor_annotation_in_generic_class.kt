// vybe-test: kotlin/kotlin_annotation_usage/test_constructor_annotation_in_generic_class
// origin: languages/kotlin/tests/kotlin/test_kotlin_annotation_usage.rs

@Target(AnnotationTarget.CONSTRUCTOR)
        annotation class Build

        class Holder<@Build T>(val value: T)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Holder(3).value).toString(), "3")
        }
