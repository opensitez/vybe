// vybe-test: kotlin/kotlin_return_expressions/test_return_last_statement_in_block
// origin: languages/kotlin/tests/kotlin/test_kotlin_return_expressions.rs

fun score(v: Int): Int {
            return {
                if (v > 1) {
                    v + 1
                } else {
                    v
                }
            }()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((score(2)).toString(), "3")
            __check((score(0)).toString(), "0")
        }
