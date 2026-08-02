// vybe-test: kotlin/in_keyword/test_in_with_if_condition
// origin: languages/kotlin/tests/kotlin/test_in_keyword.rs

fun classify(v: Int): String {
            return if (v in 1..3) "small" else if (v in 4..6) "mid" else "big"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((classify(2)).toString(), "small")
            __check((classify(5)).toString(), "mid")
            __check((classify(7)).toString(), "big")
        }
