// vybe-test: kotlin/kotlin_return_expressions/test_returning_result_from_nested_block
// origin: languages/kotlin/tests/kotlin/test_kotlin_return_expressions.rs

fun compute(v: Int): Int {
            val out = run {
                if (v < 5) {
                    return@run v * 2
                }
                v
            }
            return out
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((compute(3)).toString(), "6")
            __check((compute(7)).toString(), "7")
        }
