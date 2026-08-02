// vybe-test: kotlin/unary_ops/test_mix_increments
// origin: languages/kotlin/tests/kotlin/test_unary_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var x = 1
            var out = 0
            out += x++
            out += ++x
            __check((out).toString(), "4")
        }
