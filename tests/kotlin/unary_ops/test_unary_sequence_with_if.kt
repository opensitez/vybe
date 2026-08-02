// vybe-test: kotlin/unary_ops/test_unary_sequence_with_if
// origin: languages/kotlin/tests/kotlin/test_unary_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = -3
            __check((if (x < 0) -x else x).toString(), "3")
            __check((if (x > 0) -x else +x).toString(), "-3")
        }
