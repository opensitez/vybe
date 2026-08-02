// vybe-test: kotlin/unary_ops/test_unary_after_computation
// origin: languages/kotlin/tests/kotlin/test_unary_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var out = 0
            var x = 2
            out += +x
            out += -x
            __check((out).toString(), "0")
        }
