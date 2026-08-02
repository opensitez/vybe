// vybe-test: kotlin/when_subjects/test_when_in_keyword
// origin: languages/kotlin/tests/kotlin/test_when_subjects.rs

fun classify(ch: Char): String = when (ch) {
            'a', 'b', 'c' -> "abc"
            'x', 'y', 'z' -> "xyz"
            else -> "other"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((classify('b')).toString(), "abc")
            __check((classify('z')).toString(), "xyz")
        }
