// vybe-test: kotlin/if_expressions/test_if_with_nullable_else_branch
// origin: languages/kotlin/tests/kotlin/test_if_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: String? = null
            val out = if (value == null) "empty" else value
            __check((out).toString(), "empty")
            val other: String? = "x"
            val out2 = if (other == null) "empty" else other
            __check((out2).toString(), "x")
        }
