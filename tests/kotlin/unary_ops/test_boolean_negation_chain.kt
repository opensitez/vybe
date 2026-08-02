// vybe-test: kotlin/unary_ops/test_boolean_negation_chain
// origin: languages/kotlin/tests/kotlin/test_unary_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = !(!true)
            val b = !!false
            __check((a).toString(), "true")
            __check((b).toString(), "false")
        }
