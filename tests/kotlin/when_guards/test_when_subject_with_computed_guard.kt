// vybe-test: kotlin/when_guards/test_when_subject_with_computed_guard
// origin: languages/kotlin/tests/kotlin/test_when_guards.rs

fun kind(v: Int): String = when (v % 2) {
            0 -> if (v > 0) "even-pos" else "even-neg"
            else -> "odd"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((kind(4)).toString(), "even-pos")
            __check((kind(-3)).toString(), "odd")
        }
