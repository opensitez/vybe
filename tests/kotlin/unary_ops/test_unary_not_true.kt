// vybe-test: kotlin/unary_ops/test_unary_not_true
// origin: languages/kotlin/tests/kotlin/test_unary_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((!true).toString(), "false")
            __check((!false).toString(), "true")
        }
