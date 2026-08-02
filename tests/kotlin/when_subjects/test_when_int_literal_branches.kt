// vybe-test: kotlin/when_subjects/test_when_int_literal_branches
// origin: languages/kotlin/tests/kotlin/test_when_subjects.rs

fun valueName(x: Int): String = when (x) {
            0 -> "zero"
            1 -> "one"
            2 -> "two"
            else -> "other"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((valueName(0)).toString(), "zero")
            __check((valueName(3)).toString(), "other")
        }
