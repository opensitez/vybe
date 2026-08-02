// vybe-test: kotlin/basic/test_if_expression_used_as_return_value
// origin: languages/kotlin/tests/kotlin/test_basic.rs

fun classify(value: Int): String {
            return if (value > 10) "big" else "small"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((classify(11)).toString(), "big")
            __check((classify(3)).toString(), "small")
        }
