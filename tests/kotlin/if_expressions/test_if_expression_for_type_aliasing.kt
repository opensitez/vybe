// vybe-test: kotlin/if_expressions/test_if_expression_for_type_aliasing
// origin: languages/kotlin/tests/kotlin/test_if_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text: Any = "abc"
            val out = if (text is String) text.length else 0
            __check((out).toString(), "3")
            val num: Any = 5
            val out2 = if (num is Int) num + 1 else -1
            __check((out2).toString(), "6")
        }
