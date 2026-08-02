// vybe-test: kotlin/when_subjects/test_when_empty_else
// origin: languages/kotlin/tests/kotlin/test_when_subjects.rs

fun label(x: Int): String = when (x) {
            0 -> "zero"
            else -> "not"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((label(0)).toString(), "zero")
            __check((label(1)).toString(), "not")
        }
