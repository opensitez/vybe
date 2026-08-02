// vybe-test: kotlin/kotlin_return_expressions/test_conditional_return_to_label
// origin: languages/kotlin/tests/kotlin/test_kotlin_return_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = run outer@{
                val text = "x"
                if (text.isEmpty()) return@outer "no"
                "yes"
            }
            __check((out).toString(), "yes")
        }
