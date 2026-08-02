// vybe-test: kotlin/exceptions/test_try_expression_value_on_failure_path
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = try {
                throw Exception("boom")
            } catch (e: Exception) {
                5
            }
            __check((value).toString(), "5")
            __check(("done").toString(), "done")
        }
