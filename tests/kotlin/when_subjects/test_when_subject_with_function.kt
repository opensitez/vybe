// vybe-test: kotlin/when_subjects/test_when_subject_with_function
// origin: languages/kotlin/tests/kotlin/test_when_subjects.rs

fun isGood(x: Int): Boolean = x % 2 == 0
        fun status(x: Int): String = when (x) {
            0 -> "zero"
            else -> if (isGood(x)) "even" else "odd"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((status(4)).toString(), "even")
            __check((status(5)).toString(), "odd")
        }
