// vybe-test: kotlin/exceptions/test_try_expression_as_value_with_finally_cleanup
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = try {
                __check(("body").toString(), "body")
                9
            } finally {
                __check(("closed").toString(), "closed")
            }
            __check((value).toString(), "9")
        }
