// vybe-test: kotlin/operators/test_boolean_unary_ops
// origin: languages/kotlin/tests/kotlin/test_operators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 3
            __check((-value).toString(), "-3")
            __check((+value).toString(), "3")
            __check((!true).toString(), "false")
            __check((!false).toString(), "true")
        }
