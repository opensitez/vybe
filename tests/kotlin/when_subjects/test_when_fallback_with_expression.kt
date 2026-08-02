// vybe-test: kotlin/when_subjects/test_when_fallback_with_expression
// origin: languages/kotlin/tests/kotlin/test_when_subjects.rs

fun score(x: Int): Int = when (x) {
            in 1..3 -> 1
            in 4..6 -> 2
            else -> 3
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((score(4)).toString(), "2")
            __check((score(9)).toString(), "3")
        }
