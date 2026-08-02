// vybe-test: kotlin/kotlin_infix_keywords/test_infix_else_chain
// origin: languages/kotlin/tests/kotlin/test_kotlin_infix_keywords.rs

fun classify(v: Int): String {
            return if (v % 2 == 0 && v > 0) {
                "even"
            } else {
                "odd"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((classify(4)).toString(), "even")
            __check((classify(5)).toString(), "odd")
        }
