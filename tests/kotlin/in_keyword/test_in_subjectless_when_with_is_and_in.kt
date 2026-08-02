// vybe-test: kotlin/in_keyword/test_in_subjectless_when_with_is_and_in
// origin: languages/kotlin/tests/kotlin/test_in_keyword.rs

fun label(v: Int): String {
            return when {
                v in 1..2 -> "low"
                v in 3..4 -> "mid"
                else -> "high"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((label(2)).toString(), "low")
            __check((label(4)).toString(), "mid")
            __check((label(7)).toString(), "high")
        }
