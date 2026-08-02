// vybe-test: kotlin/kotlin_annotation_usage/test_annotation_with_long_and_char_defaults
// origin: languages/kotlin/tests/kotlin/test_kotlin_annotation_usage.rs

@Target(AnnotationTarget.FUNCTION)
        annotation class Config(val symbol: Char = 'x', val limit: Long = 10)

        @Config
        fun value() = 12

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((value()).toString(), "12")
        }
