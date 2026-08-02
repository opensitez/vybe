// vybe-test: kotlin/if_expressions/test_if_with_early_return
// origin: languages/kotlin/tests/kotlin/test_if_expressions.rs

fun classify(v: Int): Int {
            if (v < 0) return -1
            return if (v == 0) 0 else 1
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((classify(-3)).toString(), "-1")
            __check((classify(0)).toString(), "0")
            __check((classify(2)).toString(), "1")
        }
