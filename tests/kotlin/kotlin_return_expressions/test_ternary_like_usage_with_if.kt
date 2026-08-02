// vybe-test: kotlin/kotlin_return_expressions/test_ternary_like_usage_with_if
// origin: languages/kotlin/tests/kotlin/test_kotlin_return_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = if (10 % 2 == 0) "even" else "odd"
            __check((value).toString(), "even")
        }
