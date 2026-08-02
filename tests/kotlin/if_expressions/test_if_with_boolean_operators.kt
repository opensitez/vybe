// vybe-test: kotlin/if_expressions/test_if_with_boolean_operators
// origin: languages/kotlin/tests/kotlin/test_if_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = 4
            val b = 9
            val result = if (a > 2 && b > 8) "ok" else "bad"
            val second = if (a > 10 || b > 8) "yes" else "no"
            __check((result).toString(), "ok")
            __check((second).toString(), "yes")
        }
