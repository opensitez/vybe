// vybe-test: kotlin/unary_ops/test_unary_apply_chain
// origin: languages/kotlin/tests/kotlin/test_unary_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = -(+1)
            val b = (+(-2))
            __check((a).toString(), "-1")
            __check((b).toString(), "-2")
        }
