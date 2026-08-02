// vybe-test: kotlin/if_expressions/test_if_expression_assigning_immutable
// origin: languages/kotlin/tests/kotlin/test_if_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = 7
            val b = if (a > 5) a + 1 else a - 1
            val c = if (a == 10) "ten" else if (a == 7) "seven" else "other"
            __check((b).toString(), "8")
            __check((c).toString(), "seven")
        }
