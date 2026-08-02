// vybe-test: kotlin/escaped_identifiers/test_backtick_in_lambda_parameter
// origin: languages/kotlin/tests/kotlin/test_escaped_identifiers.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val fn = { `arg value`: Int -> `arg value` + 1 }
        __check((fn(6)).toString(), "7")
    }
