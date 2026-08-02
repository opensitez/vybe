// vybe-test: kotlin/kotlin_return_expressions/test_try_expression_value_with_finally
// origin: languages/kotlin/tests/kotlin/test_kotlin_return_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = try {
                "ok"
            } finally {
                __check(("final").toString(), "final")
            }
            __check((out).toString(), "ok")
        }
