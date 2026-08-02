// vybe-test: kotlin/when_subjects/test_when_returns_expression
// origin: languages/kotlin/tests/kotlin/test_when_subjects.rs

fun map(v: Int): Int = when (v) {
            1 -> 10
            2 -> 20
            else -> v
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((map(1)).toString(), "10")
            __check((map(3)).toString(), "3")
        }
