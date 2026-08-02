// vybe-test: kotlin/annotations/test_annotation_with_defaulted_constructor_argument
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

annotation class Flag(val level: String = "low")

        @Flag
        @Flag("high")
        fun run(level: String = "x"): String {
            return level
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((run()).toString(), "x")
        }
