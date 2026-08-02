// vybe-test: kotlin/when_expressions/test_when_with_char_subject
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

fun kind(ch: Char): String {
            return when (ch) {
                in 'a'..'f' -> "low"
                in 'g'..'m' -> "mid"
                in 'n'..'z' -> "high"
                else -> "other"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((kind('c')).toString(), "low")
            __check((kind('h')).toString(), "mid")
            __check((kind('x')).toString(), "high")
            __check((kind('2')).toString(), "other")
        }
