// vybe-test: kotlin/when_expressions/test_when_with_subject_in_function_reference_style
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

fun classify(value: Int): String {
            val fn = { n: Int ->
                when (n) {
                    1 -> "single"
                    2, 3 -> "pair"
                    in 4..6 -> "few"
                    else -> "many"
                }
            }
            return fn(value)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((classify(1)).toString(), "single")
            __check((classify(3)).toString(), "pair")
            __check((classify(5)).toString(), "few")
            __check((classify(9)).toString(), "many")
        }
