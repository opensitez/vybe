// vybe-test: kotlin/unary_ops/test_not_chain
// origin: languages/kotlin/tests/kotlin/test_unary_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((!!false).toString(), "true")
            __check((!true).toString(), "false")
        }
