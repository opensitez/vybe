// vybe-test: kotlin/if_expressions/test_if_chain_with_elif
// origin: languages/kotlin/tests/kotlin/test_if_expressions.rs

fun classify(v: Int): String {
            return if (v < 0) "neg" else if (v == 0) "zero" else if (v in 1..10) "small" else "big"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((classify(-1)).toString(), "neg")
            __check((classify(0)).toString(), "zero")
            __check((classify(7)).toString(), "small")
            __check((classify(15)).toString(), "big")
        }
